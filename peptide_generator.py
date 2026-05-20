#!/usr/bin/env python3
"""
Peptide Take Home Generator
Reads Excel pricing data and generates a customized Word document
for patient take-home instructions.

Usage (command line):
    python peptide_generator.py --medication "BPC157" --month "Month 1" \
        --patient "Jane Smith" --appointment "6/15/2026"

Usage (import):
    from peptide_generator import read_workbook, generate_document
"""

import os
import re
import sys
import copy
import shutil
import argparse
from pathlib import Path

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from lxml import etree

# ============================================================
# PATHS  (relative to this script's directory)
# ============================================================

BASE_DIR = Path(__file__).parent
WORKBOOK_PATH   = BASE_DIR / 'FINAL FINAL Pricing 100 per bottle.xlsx'
TEMPLATE_PATH   = BASE_DIR / 'Coding Take Home Instrucitons 3ml.docx'
OUTPUT_DIR      = BASE_DIR / 'Generated'

INFO_SHEETS = {
    'Tesam 10':  BASE_DIR / 'Tesamorelin Flyer.docx',
    'Keith Tesa': BASE_DIR / 'Tesamorelin Flyer.docx',
    'BPC+TB500': BASE_DIR / 'BPC TB Flyer.docx',
}

IGNORED_SHEETS = {
    'Patient Prices', 'Pep Costs', 'FAQ',
    'Whats it for', 'COSTS', 'Marketing', 'Discount',
}

# ============================================================
# HELPERS
# ============================================================

def clean_number(value):
    """Return a value as a clean integer or decimal string."""
    if value is None or str(value).strip() == '':
        return ''
    try:
        num = float(str(value))
        if abs(num - round(num)) < 0.000001:
            return str(int(round(num)))
        return f'{num:.4f}'.rstrip('0').rstrip('.')
    except (ValueError, TypeError):
        return str(value)


def to_currency(value):
    """Format a value as USD currency string."""
    if value is None or str(value).strip() == '':
        return ''
    clean = str(value).strip().replace('$', '').replace(',', '')
    try:
        return f'${float(clean):,.2f}'
    except (ValueError, TypeError):
        return str(value)


def safe_filename(value):
    """Strip characters that are illegal in Windows file names."""
    return re.sub(r'[\\/:*?"<>|]', '-', str(value))


# ============================================================
# EXCEL PARSING
# ============================================================

def _sheet_rows(ws):
    """Return {row_number: {col_letter: value}} for every non-blank row."""
    rows = {}
    for row in ws.iter_rows():
        row_num = row[0].row
        row_data = {}
        for cell in row:
            # Skip MergedCell objects which lack column_letter
            try:
                row_data[cell.column_letter] = cell.value
            except AttributeError:
                pass
        rows[row_num] = row_data
    return rows


def _parse_month_block(rows, start_row, sheet_name, label_override=None):
    """
    Parse one Month / CUSTOM PROTOCOL block.
    Returns a dict or None if no usable data is found.
    """
    header    = rows.get(start_row, {})
    month_label = label_override or str(header.get('A', '')).strip()
    if not month_label:
        return None

    # ---- week rows (up to 8 rows below start) ----
    week_rows     = []
    last_week_row = start_row
    for r in range(start_row + 1, start_row + 9):
        if r not in rows:
            continue
        rd = rows[r]
        wl = rd.get('A')
        if wl and re.match(r'^(Week|week)', str(wl)):
            label = str(wl)
            if re.match(r'^(Week|week)$', label):
                label = f'Week {len(week_rows) + 1}'
            week_rows.append({
                'week':               label,
                'frequency':          str(rd.get('B') or ''),
                'mcg':                str(rd.get('C') or ''),
                'units_per_injection': clean_number(rd.get('D')),
                'injections':         clean_number(rd.get('E')),
                'units_needed':       clean_number(rd.get('F')),
            })
            last_week_row = r

    # ---- summary row: three fallback strategies ----
    summary_row = None
    # 1. No A value, has F and K
    for r in range(last_week_row + 1, last_week_row + 5):
        if r not in rows:
            continue
        rd = rows[r]
        if not rd.get('A') and rd.get('F') and rd.get('K'):
            summary_row = rd
            break
    # 2. Column M = 'Profit'
    if not summary_row:
        for r in range(start_row + 1, start_row + 11):
            if r not in rows:
                continue
            rd = rows[r]
            if str(rd.get('M') or '').strip() == 'Profit':
                summary_row = rd
                break
    # 3. Has J and K but not A
    if not summary_row:
        for r in range(start_row + 1, start_row + 11):
            if r not in rows:
                continue
            rd = rows[r]
            if rd.get('J') and rd.get('K') and not rd.get('A'):
                summary_row = rd
                break

    # ---- cartridge count ----
    cartridge_count = ''
    for r in range(start_row + 1, start_row + 13):
        if r not in rows:
            continue
        rd = rows[r]
        if 'Number of Cartridges needed' in str(rd.get('E') or ''):
            cartridge_count = clean_number(rd.get('F'))
            break

    # ---- supply counts (columns H and I) ----
    supplies = {
        'Pen': '', 'Syringes': '', 'Pen Needles': '',
        'Cartridges': '', 'Alcohol Prep Pads': '', 'Storage Case': '',
    }
    for r in range(start_row, start_row + 9):
        if r not in rows:
            continue
        rd      = rows[r]
        item    = str(rd.get('I') or '')
        count   = clean_number(rd.get('H'))
        if not item:
            continue
        if item == 'Pen':
            supplies['Pen'] = count
        elif re.search(r'Syringe', item, re.I):
            supplies['Syringes'] = count
        elif re.search(r'Pen Needles', item, re.I):
            supplies['Pen Needles'] = count
        elif re.search(r'Cartridges', item, re.I):
            supplies['Cartridges'] = count
        elif re.search(r'Alochol|Alcohol', item, re.I):
            supplies['Alcohol Prep Pads'] = count
        elif re.search(r'Case', item, re.I):
            supplies['Storage Case'] = count

    if cartridge_count:
        supplies['Cartridges'] = cartridge_count

    total_units   = clean_number(summary_row.get('F')) if summary_row else clean_number(header.get('H'))
    patient_price = to_currency(summary_row.get('K'))  if summary_row else to_currency(header.get('K'))
    frequencies   = list(dict.fromkeys(w['frequency'] for w in week_rows if w['frequency']))

    return {
        'label':         month_label,
        'sheet_name':    sheet_name,
        'weeks':         week_rows,
        'frequency':     ', '.join(frequencies),
        'total_units':   total_units,
        'patient_price': patient_price,
        'supplies':      supplies,
        'is_custom':     bool(label_override),
    }


def read_workbook(path=None):
    """
    Read the Excel workbook and return an ordered dict:
        { sheet_name: { 'display_name', 'sheet_name', 'months': [...] } }
    """
    path = Path(path) if path else WORKBOOK_PATH
    wb   = openpyxl.load_workbook(str(path), data_only=True)
    meds = {}

    for sheet_name in wb.sheetnames:
        if sheet_name in IGNORED_SHEETS:
            continue
        ws   = wb[sheet_name]
        rows = _sheet_rows(ws)

        display_name = rows.get(1, {}).get('A') or sheet_name
        month_blocks = []

        for rn in sorted(rows.keys()):
            label = rows[rn].get('A')
            if label is None:
                continue
            ls = str(label).strip()
            if re.match(r'^Month\s+\d+', ls):
                blk = _parse_month_block(rows, rn, sheet_name)
                if blk:
                    month_blocks.append(blk)
            elif ls == 'CUSTOM PROTOCOL':
                blk = _parse_month_block(rows, rn, sheet_name, label_override='Custom Protocol')
                if blk:
                    month_blocks.append(blk)

        if month_blocks:
            meds[sheet_name] = {
                'sheet_name':   sheet_name,
                'display_name': str(display_name),
                'months':       month_blocks,
            }

    return meds


# ============================================================
# WORD XML HELPERS
# ============================================================

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _para_text_key(para):
    """Return the paragraph's joined run text, stripped of tabs/spaces."""
    return para.text.replace('\t', '').strip()


def _clear_runs(para):
    """Remove all w:r children from a paragraph element."""
    p = para._p
    for r in p.findall(qn('w:r')):
        p.remove(r)


def _set_split_text(para, left_text, right_text, font_pt=12):
    """
    Replace paragraph content with:
        <left_text>  [right-tab at 9360 twips]  <right_text>
    Mimics Set-ParagraphSplitText from the PowerShell script.
    """
    p = para._p

    # Ensure pPr exists
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p.insert(0, pPr)

    # Remove old tab stops, add right tab at 9360
    for old in pPr.findall(qn('w:tabs')):
        pPr.remove(old)
    tabs    = OxmlElement('w:tabs')
    rt      = OxmlElement('w:tab')
    rt.set(qn('w:val'), 'right')
    rt.set(qn('w:pos'), '9360')
    tabs.append(rt)
    pPr.append(tabs)

    _clear_runs(para)

    def _make_run(text):
        r  = OxmlElement('w:r')
        rp = OxmlElement('w:rPr')
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), str(int(font_pt * 2)))
        rp.append(sz)
        r.append(rp)
        t  = OxmlElement('w:t')
        t.text = text
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        return r

    def _make_tab_run():
        r  = OxmlElement('w:r')
        r.append(OxmlElement('w:tab'))
        return r

    p.append(_make_run(left_text))
    p.append(_make_tab_run())
    p.append(_make_run(right_text))


def _insert_right_para_after(para, text, font_pt=12):
    """
    Insert a new right-aligned paragraph immediately after *para*.
    Mimics Add-RightAlignedParagraphAfter from the PowerShell script.
    """
    new_p  = OxmlElement('w:p')
    pPr    = OxmlElement('w:pPr')
    jc     = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'right')
    pPr.append(jc)
    new_p.append(pPr)

    r  = OxmlElement('w:r')
    rp = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_pt * 2)))
    rp.append(sz)
    r.append(rp)
    t  = OxmlElement('w:t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    new_p.append(r)

    para._p.addnext(new_p)


def _set_cell(cell, text, bold=False, center=True, font_pt=None):
    """Clear a table cell and write *text* into it."""
    # Wipe existing paragraphs except the first (preserve cell formatting)
    tc = cell._tc
    for extra_p in tc.findall(qn('w:p'))[1:]:
        tc.remove(extra_p)

    para = cell.paragraphs[0]
    # Clear existing runs
    for r in para._p.findall(qn('w:r')):
        para._p.remove(r)

    # Alignment
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        para._p.insert(0, pPr)
    for old_jc in pPr.findall(qn('w:jc')):
        pPr.remove(old_jc)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center' if center else 'left')
    pPr.append(jc)

    run = para.add_run(str(text))
    if bold:
        run.bold = True
    if font_pt:
        run.font.size = Pt(font_pt)


def _shade_cell(cell, fill_hex):
    """Set cell background shading."""
    tc   = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill_hex)
    tcPr.append(shd)


def _page_break_element():
    """Return a <w:p> containing a page break."""
    p  = OxmlElement('w:p')
    r  = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    return p


def _insert_page_break_before(table):
    """Insert a page-break paragraph immediately before *table*."""
    table._tbl.addprevious(_page_break_element())


def _clear_table_data_rows(table):
    """Remove all rows after the header row (row 0)."""
    tbl = table._tbl
    for tr in list(tbl.findall(qn('w:tr')))[1:]:
        tbl.remove(tr)


def _append_table_row(table, texts, bold=False, center=True,
                      font_pt=None, shade_hex=None):
    """Append a new data row copied from the header row structure."""
    new_tr = copy.deepcopy(table.rows[0]._tr)
    table._tbl.append(new_tr)
    new_row = table.rows[-1]
    for i, text in enumerate(texts):
        if i < len(new_row.cells):
            _set_cell(new_row.cells[i], str(text), bold=bold,
                      center=center, font_pt=font_pt)
            if shade_hex:
                _shade_cell(new_row.cells[i], shade_hex)


def _insert_blank_para_after_table(table):
    """Insert an empty paragraph immediately after the table."""
    table._tbl.addnext(OxmlElement('w:p'))


def _merge_flyer_before_table(doc, table_index, flyer_path):
    """
    Copy every element from *flyer_path*'s body and insert it
    (preceded by a page break) immediately before the given table.
    """
    flyer_path = Path(flyer_path)
    if not flyer_path.exists():
        return

    flyer_body = Document(str(flyer_path)).element.body
    ref_table  = doc.tables[table_index]._tbl

    # Insert a page break first (will appear after existing content)
    pb = _page_break_element()
    ref_table.addprevious(pb)

    # Insert flyer body elements before the page break
    for elem in list(flyer_body):
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'sectPr':
            continue
        pb.addprevious(copy.deepcopy(elem))


# ============================================================
# DOCUMENT GENERATION
# ============================================================

def generate_document(
    med_data:    dict,
    month_data:  dict,
    patient_name: str = '',
    appointment:  str = '',
    output_path:  Path = None,
    workbook_path: Path = None,
) -> Path:
    """
    Generate a filled take-home Word document.

    Parameters
    ----------
    med_data      : one entry from read_workbook()
    month_data    : one month block from med_data['months']
    patient_name  : optional patient name string
    appointment   : optional next-appointment string
    output_path   : explicit save path; auto-generated if omitted
    workbook_path : path to the Excel file (for the output folder)

    Returns
    -------
    Path to the generated .docx file
    """
    OUTPUT_DIR.mkdir(exist_ok=True)

    if output_path is None:
        filename = (
            f"{safe_filename(patient_name or 'Patient')} - "
            f"{safe_filename(med_data['display_name'])} - "
            f"{safe_filename(month_data['label'])}.docx"
        )
        output_path = OUTPUT_DIR / filename

    # Work on a fresh copy of the template
    shutil.copy2(str(TEMPLATE_PATH), str(output_path))
    doc = Document(str(output_path))

    # ---- Paragraph substitutions ----
    paras_to_delete = []
    for para in doc.paragraphs:
        key = _para_text_key(para)
        if key == 'MedicationPatient Name':
            _set_split_text(para, med_data['display_name'], patient_name, font_pt=18)

        elif key == 'FrequencyNext Appt:':
            _set_split_text(para, f"Frequency: {month_data['frequency']}", 'Next Appointment', font_pt=12)
            _insert_right_para_after(para, appointment, font_pt=12)

        elif key == 'Frequency':
            _set_split_text(para, f"Frequency: {month_data['frequency']}", 'Next Appointment', font_pt=12)
            _insert_right_para_after(para, appointment, font_pt=12)

        elif key == '____________':
            paras_to_delete.append(para)

        elif key == 'Month 1: week 1-4':
            week_labels = ', '.join(w['week'] for w in month_data['weeks'])
            _clear_runs(para)
            para.add_run(f"{month_data['label']}: {week_labels}")

    for para in paras_to_delete:
        para._p.getparent().remove(para._p)

    # ---- Page break before the injection-log table (Table 1) ----
    _insert_page_break_before(doc.tables[1])

    # ---- Dosing table (Table 0) — dynamic rows ----
    dose_tbl = doc.tables[0]
    _clear_table_data_rows(dose_tbl)

    if month_data['is_custom']:
        # Custom protocol: Week | Frequency | Units per Injection | # of Injections
        hdr_cells = dose_tbl.rows[0].cells
        _set_cell(hdr_cells[0], 'Week',                 bold=True)
        _set_cell(hdr_cells[1], 'Frequency',            bold=True)
        if len(hdr_cells) > 2:
            _set_cell(hdr_cells[2], 'Units per Injection', bold=True)
        if len(hdr_cells) > 3:
            _set_cell(hdr_cells[3], '# of Injections',     bold=True)

        for wk in month_data['weeks']:
            row_data = [wk['week'], wk['frequency'],
                        wk['units_per_injection'], wk['injections']]
            _append_table_row(dose_tbl, row_data, shade_hex='DDEBF7')

        # Shade header row too
        for cell in dose_tbl.rows[0].cells:
            _shade_cell(cell, 'DDEBF7')

    else:
        # Standard layout: Week | Frequency | Units to Inject
        hdr_cells = dose_tbl.rows[0].cells
        _set_cell(hdr_cells[0], month_data['label'], bold=True)
        if len(hdr_cells) > 1:
            _set_cell(hdr_cells[1], 'Frequency',     bold=True)
        if len(hdr_cells) > 2:
            _set_cell(hdr_cells[2], 'Inject',        bold=True)

        for wk in month_data['weeks']:
            row_data = [wk['week'], wk['frequency'],
                        f"{wk['units_per_injection']} units"]
            _append_table_row(dose_tbl, row_data)

    # Blank line after dosing table for readability
    _insert_blank_para_after_table(dose_tbl)

    # ---- Pricing table (Table 2) ----
    pt   = doc.tables[2]
    wks  = ', '.join(w['week'] for w in month_data['weeks'])
    sup  = month_data['supplies']

    _set_cell(pt.rows[0].cells[0], f"{month_data['label']}: {wks}",       center=False)
    _set_cell(pt.rows[0].cells[3], 'Price',                                center=False)
    _set_cell(pt.rows[1].cells[0], med_data['display_name'],               center=False)
    _set_cell(pt.rows[1].cells[1], f"{month_data['total_units']} units",   center=False)
    _set_cell(pt.rows[1].cells[3], month_data['patient_price'],            center=False)
    _set_cell(pt.rows[2].cells[0], 'Cartridge',                            center=False)
    _set_cell(pt.rows[2].cells[1], sup['Cartridges'],                      center=False)

    if sup['Syringes']:
        _set_cell(pt.rows[3].cells[0], 'Syringes',    center=False)
        _set_cell(pt.rows[3].cells[1], sup['Syringes'], center=False)
    else:
        _set_cell(pt.rows[3].cells[0], 'Pen Needles',    center=False)
        _set_cell(pt.rows[3].cells[1], sup['Pen Needles'], center=False)

    _set_cell(pt.rows[4].cells[0], 'Alcohol Prep Pads',        center=False)
    _set_cell(pt.rows[4].cells[1], sup['Alcohol Prep Pads'],   center=False)

    if len(pt.rows) > 6:
        _set_cell(pt.rows[6].cells[0], 'Storage Case',     center=False)
        _set_cell(pt.rows[6].cells[1], sup['Storage Case'], center=False)
    if len(pt.rows) > 7:
        _set_cell(pt.rows[7].cells[0], 'Peptide Pen (includes case)', center=False)
        _set_cell(pt.rows[7].cells[1], sup['Pen'],                     center=False)

    # ---- Total row: blank spacer + bold large total ----
    if month_data.get('patient_price'):
        num_cols = len(pt.rows[0].cells)
        # Blank spacer row
        _append_table_row(pt, [''] * num_cols)
        # Total row — label in first col, price in last col, larger font
        total_data = [''] * num_cols
        total_data[0]        = 'Total'
        total_data[num_cols - 1] = month_data['patient_price']
        _append_table_row(pt, total_data, bold=True, center=False, font_pt=14)

    # ---- Optional info-sheet flyer (Tesamorelin, BPC+TB500) ----
    flyer = INFO_SHEETS.get(med_data['sheet_name'])
    if flyer and Path(flyer).exists():
        _merge_flyer_before_table(doc, 2, flyer)

    doc.save(str(output_path))
    return output_path


# ============================================================
# CLI ENTRY POINT
# ============================================================

def main():
    parser = argparse.ArgumentParser(description='Peptide Take Home Generator')
    parser.add_argument('--medication',      required=False, help='Sheet name, e.g. "BPC157"')
    parser.add_argument('--month',           required=False, help='Month label, e.g. "Month 1"')
    parser.add_argument('--patient',         default='',     help='Patient name (optional)')
    parser.add_argument('--appointment',     default='',     help='Next appointment date (optional)')
    parser.add_argument('--workbook',        default=None,   help='Path to Excel workbook')
    parser.add_argument('--list',            action='store_true', help='List all available medications and months')
    args = parser.parse_args()

    wb_path = Path(args.workbook) if args.workbook else WORKBOOK_PATH
    data    = read_workbook(wb_path)

    if args.list or not (args.medication and args.month):
        print('\nAvailable medications and months:\n')
        for sheet, info in data.items():
            months = ', '.join(m['label'] for m in info['months'])
            print(f"  {sheet:20s}  →  {months}")
        if not (args.medication and args.month):
            return

    if args.medication not in data:
        print(f"ERROR: Medication '{args.medication}' not found.")
        print(f"Available: {', '.join(data.keys())}")
        sys.exit(1)

    med_data   = data[args.medication]
    month_data = next(
        (m for m in med_data['months'] if m['label'] == args.month), None
    )
    if not month_data:
        available = ', '.join(m['label'] for m in med_data['months'])
        print(f"ERROR: Month '{args.month}' not found. Available: {available}")
        sys.exit(1)

    out = generate_document(
        med_data     = med_data,
        month_data   = month_data,
        patient_name = args.patient,
        appointment  = args.appointment,
    )
    print(f"Generated: {out}")


if __name__ == '__main__':
    main()