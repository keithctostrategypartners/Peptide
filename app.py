#!/usr/bin/env python3
"""
Peptide Take Home Generator - Cloud-Ready Web App
Run locally:  python app.py  then open http://localhost:5000

Environment variables (set these on Railway / your host):
  SECRET_KEY      random string for session security
  ADMIN_PASSWORD  password for the Excel-upload admin panel
  STAFF_USERS     JSON like {"front_desk":"pass1","nurse":"pass2"}
"""

import os, io, json, shutil, tempfile
from pathlib import Path
from datetime import datetime
from functools import wraps

from flask import (Flask, request, jsonify, send_file,
                   redirect, url_for, session, render_template_string)
from peptide_generator import read_workbook, generate_document, WORKBOOK_PATH, safe_filename

SECRET_KEY     = os.environ.get('SECRET_KEY', 'dev-secret-change-in-production')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'peptide2026!')

_staff_env = os.environ.get('STAFF_USERS', '')
try:
    STAFF_USERS = json.loads(_staff_env) if _staff_env else {"staff": "peptide2026!"}
except json.JSONDecodeError:
    STAFF_USERS = {"staff": "peptide2026!"}

app = Flask(__name__)
app.secret_key = SECRET_KEY

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated

LOGIN_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sign In</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
     background:#f0f4f8;min-height:100vh;display:flex;
     align-items:center;justify-content:center}
.card{background:#fff;border-radius:12px;
      box-shadow:0 2px 16px rgba(0,0,0,.1);
      padding:40px 44px;width:100%;max-width:380px}
h1{font-size:1.3rem;font-weight:700;color:#1a2a3a;margin-bottom:4px}
.sub{font-size:.85rem;color:#6b7a8d;margin-bottom:28px}
label{display:block;font-size:.82rem;font-weight:600;color:#374151;margin-bottom:5px}
input{width:100%;padding:10px 12px;border:1.5px solid #d1d9e0;
      border-radius:7px;font-size:.95rem;background:#fafbfc;margin-bottom:16px}
input:focus{outline:none;border-color:#4f7ef0;background:#fff}
.btn{width:100%;padding:12px;background:#4f7ef0;color:#fff;
     border:none;border-radius:8px;font-size:1rem;font-weight:600;cursor:pointer}
.btn:hover{background:#3a67d4}
.error{background:#fdf0f0;color:#8b1a1a;border:1px solid #f0a9a9;
       border-radius:7px;padding:10px 14px;font-size:.88rem;margin-bottom:16px}
</style></head><body>
<div class="card">
  <h1>&#128138; Peptide Generator</h1>
  <p class="sub">Sign in to continue</p>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Username</label>
    <input type="text" name="username" autocomplete="username" required>
    <label>Password</label>
    <input type="password" name="password" autocomplete="current-password" required>
    <button class="btn" type="submit">Sign In</button>
  </form>
</div></body></html>"""

MAIN_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Peptide Take Home Generator</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
     background:#f0f4f8;min-height:100vh;
     display:flex;flex-direction:column;align-items:center;padding:40px 16px}
.topbar{width:100%;max-width:560px;display:flex;justify-content:flex-end;
        align-items:center;margin-bottom:12px;gap:12px}
.topbar span{font-size:.83rem;color:#6b7a8d}
.topbar a{font-size:.83rem;color:#4f7ef0;text-decoration:none}
.topbar a:hover{text-decoration:underline}
.card{background:#fff;border-radius:12px;
      box-shadow:0 2px 16px rgba(0,0,0,.1);
      padding:36px 40px;width:100%;max-width:560px;margin-bottom:24px}
h1{font-size:1.45rem;font-weight:700;color:#1a2a3a;margin-bottom:6px}
.subtitle{font-size:.88rem;color:#6b7a8d;margin-bottom:28px}
label{display:block;font-size:.82rem;font-weight:600;color:#374151;margin-bottom:5px}
select,input[type=text],input[type=password]{
  width:100%;padding:10px 12px;border:1.5px solid #d1d9e0;
  border-radius:7px;font-size:.95rem;color:#1a2a3a;
  background:#fafbfc;transition:border-color .15s;appearance:none}
select:focus,input:focus{outline:none;border-color:#4f7ef0;background:#fff}
.field{margin-bottom:18px}
.row{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.btn{width:100%;padding:12px;background:#4f7ef0;color:#fff;
     border:none;border-radius:8px;font-size:1rem;font-weight:600;
     cursor:pointer;transition:background .15s;margin-top:4px}
.btn:hover{background:#3a67d4}
.btn:disabled{background:#9cb3f0;cursor:not-allowed}
#status{margin-top:18px;padding:12px 16px;border-radius:8px;
        font-size:.9rem;display:none}
#status.success{background:#eaf7ee;color:#186a35;border:1px solid #a3d9b1}
#status.error{background:#fdf0f0;color:#8b1a1a;border:1px solid #f0a9a9}
#status.info{background:#eef3ff;color:#1a3a8b;border:1px solid #a9bef0}
.dl-btn{display:none;margin-top:12px;width:100%;padding:12px;
        background:#186a35;color:#fff;border:none;border-radius:8px;
        font-size:1rem;font-weight:600;cursor:pointer;text-align:center;
        text-decoration:none}
.dl-btn:hover{background:#145529}
hr{border:none;border-top:1px solid #e5e9f0;margin:20px 0}
.admin-toggle{font-size:.8rem;color:#9ca3af;cursor:pointer;
              text-align:center;user-select:none}
.admin-toggle:hover{color:#6b7280}
.admin-card{display:none}
.admin-card.open{display:block}
.admin-header{display:flex;align-items:center;gap:8px;margin-bottom:18px}
.admin-header h2{font-size:1.1rem;font-weight:700;color:#7c3aed}
.badge{background:#ede9fe;color:#5b21b6;font-size:.72rem;
       font-weight:700;padding:2px 8px;border-radius:20px}
.file-info{font-size:.82rem;color:#6b7a8d;margin-top:6px}
#adminStatus{margin-top:14px;padding:10px 14px;border-radius:7px;
             font-size:.88rem;display:none}
#adminStatus.success{background:#eaf7ee;color:#186a35;border:1px solid #a3d9b1}
#adminStatus.error{background:#fdf0f0;color:#8b1a1a;border:1px solid #f0a9a9}
#adminStatus.info{background:#eef3ff;color:#1a3a8b;border:1px solid #a9bef0}
</style></head><body>

<div class="topbar">
  <span>Signed in as <strong>{{ username }}</strong></span>
  <a href="/logout">Sign out</a>
</div>

<div class="card">
  <h1>&#128138; Peptide Take Home Generator</h1>
  <p class="subtitle">Select a medication and month, then generate the patient document.</p>
  <div class="field">
    <label>Medication</label>
    <select id="medication"><option value="">&#8212; Loading&#8230; &#8212;</option></select>
  </div>
  <div class="field">
    <label>Month / Protocol</label>
    <select id="month" disabled><option value="">&#8212; Choose medication first &#8212;</option></select>
  </div>
  <div class="row">
    <div class="field">
      <label>Patient Name <span style="font-weight:400;color:#9ca3af">(optional)</span></label>
      <input type="text" id="patient" placeholder="e.g. Jane Smith">
    </div>
    <div class="field">
      <label>Next Appointment <span style="font-weight:400;color:#9ca3af">(optional)</span></label>
      <input type="text" id="appt" placeholder="e.g. 6/15/2026">
    </div>
  </div>
  <button class="btn" id="generateBtn" onclick="generateDoc()">Generate Word Document</button>
  <div id="status"></div>
  <a class="dl-btn" id="dlBtn" href="#" download>&#11015; Download Document</a>
</div>

<div class="admin-toggle" onclick="toggleAdmin()">&#9881; Admin</div>

<div class="card admin-card" id="adminCard">
  <div class="admin-header">
    <span>&#128272;</span><h2>Admin Panel</h2><span class="badge">Restricted</span>
  </div>
  <p style="font-size:.85rem;color:#6b7a8d;margin-bottom:18px">
    Upload a new Excel pricing file. The current file is backed up automatically.
  </p>
  <div class="field">
    <label>Admin Password</label>
    <input type="password" id="adminPass" placeholder="Enter admin password">
  </div>
  <hr>
  <div class="field">
    <label>New Excel File (.xlsx)</label>
    <input type="file" id="xlsxFile" accept=".xlsx"
           style="padding:8px;background:#f8f8ff" onchange="showFileName()">
    <div class="file-info" id="fileInfo">No file chosen</div>
  </div>
  <button class="btn" onclick="uploadExcel()">Upload &amp; Replace Pricing File</button>
  <div id="adminStatus"></div>
</div>

<script>
let medData = {};

async function loadMedications() {
  try {
    const r = await fetch('/api/medications');
    if (r.status === 401) { location = '/login'; return; }
    medData = await r.json();
    const sel = document.getElementById('medication');
    sel.innerHTML = '<option value="">&#8212; Select medication &#8212;</option>';
    for (const k of Object.keys(medData)) {
      const o = document.createElement('option');
      o.value = k; o.textContent = medData[k].display_name;
      sel.appendChild(o);
    }
  } catch(e) { setStatus('error','Could not load medications.'); }
}

document.getElementById('medication').addEventListener('change', function() {
  const ms = document.getElementById('month');
  ms.innerHTML = '';
  if (!this.value || !medData[this.value]) {
    ms.innerHTML = '<option value="">&#8212; Choose medication first &#8212;</option>';
    ms.disabled = true; return;
  }
  for (const m of medData[this.value].months) {
    const o = document.createElement('option');
    o.value = o.textContent = m.label; ms.appendChild(o);
  }
  ms.disabled = false;
});

async function generateDoc() {
  const med = document.getElementById('medication').value;
  const mon = document.getElementById('month').value;
  if (!med || !mon) { setStatus('error','Please select a medication and month.'); return; }
  const btn = document.getElementById('generateBtn');
  btn.disabled = true; btn.textContent = 'Generating…';
  document.getElementById('dlBtn').style.display = 'none';
  setStatus('info','Generating document…');
  try {
    const r = await fetch('/api/generate', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({
        medication: med, month: mon,
        patient_name: document.getElementById('patient').value.trim(),
        appointment:  document.getElementById('appt').value.trim()
      })
    });
    if (r.status === 401) { location = '/login'; return; }
    const d = await r.json();
    if (r.ok) {
      setStatus('success','✓ Document ready — click below to download.');
      const dl = document.getElementById('dlBtn');
      dl.href = '/download/' + encodeURIComponent(d.filename);
      dl.download = d.filename;
      dl.style.display = 'block';
    } else { setStatus('error','Error: ' + d.error); }
  } catch(e) { setStatus('error','Network error. Please try again.'); }
  finally { btn.disabled = false; btn.textContent = 'Generate Word Document'; }
}

function toggleAdmin() { document.getElementById('adminCard').classList.toggle('open'); }
function showFileName() {
  const f = document.getElementById('xlsxFile').files[0];
  document.getElementById('fileInfo').textContent = f ? f.name : 'No file chosen';
}

async function uploadExcel() {
  const pass = document.getElementById('adminPass').value;
  const file = document.getElementById('xlsxFile').files[0];
  if (!pass) { setAdminStatus('error','Enter the admin password.'); return; }
  if (!file)  { setAdminStatus('error','Choose a .xlsx file.'); return; }
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    setAdminStatus('error','File must be .xlsx'); return; }
  const form = new FormData();
  form.append('password', pass); form.append('file', file);
  setAdminStatus('info','Uploading…');
  try {
    const r = await fetch('/admin/upload', {method:'POST', body:form});
    if (r.status === 401) { location = '/login'; return; }
    const d = await r.json();
    if (r.ok) {
      setAdminStatus('success','✓ ' + d.message + ' Reloading list…');
      document.getElementById('xlsxFile').value = '';
      document.getElementById('fileInfo').textContent = 'No file chosen';
      setTimeout(() => loadMedications(), 1200);
    } else { setAdminStatus('error','Error: ' + d.error); }
  } catch(e) { setAdminStatus('error','Upload failed.'); }
}

function setStatus(t,m){const e=document.getElementById('status');e.className=t;e.textContent=m;e.style.display='block';}
function setAdminStatus(t,m){const e=document.getElementById('adminStatus');e.className=t;e.textContent=m;e.style.display='block';}
loadMedications();
</script></body></html>"""

@app.route('/login', methods=['GET','POST'])
def login():
    error = None
    if request.method == 'POST':
        u = request.form.get('username','').strip()
        p = request.form.get('password','')
        if STAFF_USERS.get(u) == p:
            session['logged_in'] = True
            session['username']  = u
            return redirect(request.args.get('next') or '/')
        error = 'Incorrect username or password.'
    return render_template_string(LOGIN_HTML, error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

@app.route('/')
@login_required
def index():
    return render_template_string(MAIN_HTML, username=session.get('username','staff'))

@app.route('/api/medications')
@login_required
def api_medications():
    try:
        data = read_workbook(WORKBOOK_PATH)
        return jsonify({
            k: {'display_name': v['display_name'],
                'months': [{'label': m['label']} for m in v['months']]}
            for k, v in data.items()
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate', methods=['POST'])
@login_required
def api_generate():
    body        = request.get_json(force=True)
    med_key     = body.get('medication','').strip()
    month_label = body.get('month','').strip()
    patient     = body.get('patient_name','').strip()
    appt        = body.get('appointment','').strip()

    if not med_key or not month_label:
        return jsonify({'error': 'medication and month are required'}), 400

    try:
        data = read_workbook(WORKBOOK_PATH)
    except Exception as e:
        return jsonify({'error': f'Could not read Excel file: {e}'}), 500

    if med_key not in data:
        return jsonify({'error': f'Medication "{med_key}" not found'}), 404

    med_data   = data[med_key]
    month_data = next((m for m in med_data['months'] if m['label'] == month_label), None)
    if not month_data:
        return jsonify({'error': f'Month "{month_label}" not found'}), 404

    try:
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            tmp_path = Path(tmp.name)

        generate_document(
            med_data=med_data, month_data=month_data,
            patient_name=patient, appointment=appt,
            output_path=tmp_path,
        )

        filename = (
            f"{safe_filename(patient or 'Patient')} - "
            f"{safe_filename(med_data['display_name'])} - "
            f"{safe_filename(month_data['label'])}.docx"
        )
        file_bytes = tmp_path.read_bytes()
        tmp_path.unlink(missing_ok=True)

        # Hold generated file in memory keyed by filename
        app.config.setdefault('_docs', {})[filename] = file_bytes
        return jsonify({'filename': filename})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
@login_required
def download(filename):
    data = app.config.get('_docs', {}).get(filename)
    if not data:
        return 'File not found or expired — please generate again.', 404
    return send_file(
        io.BytesIO(data),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=filename,
    )

@app.route('/admin/upload', methods=['POST'])
@login_required
def admin_upload():
    if request.form.get('password','') != ADMIN_PASSWORD:
        return jsonify({'error': 'Incorrect admin password'}), 403

    file = request.files.get('file')
    if not file or not file.filename:
        return jsonify({'error': 'No file received'}), 400
    if not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'File must be a .xlsx Excel file'}), 400

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
        file.save(tmp_path)

    try:
        import openpyxl
        openpyxl.load_workbook(tmp_path, data_only=True)
    except Exception as e:
        os.unlink(tmp_path)
        return jsonify({'error': f'Invalid Excel file: {e}'}), 400

    if WORKBOOK_PATH.exists():
        ts     = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup = WORKBOOK_PATH.with_name(f'BACKUP_{ts}_{WORKBOOK_PATH.name}')
        shutil.copy2(str(WORKBOOK_PATH), str(backup))

    shutil.move(tmp_path, str(WORKBOOK_PATH))
    return jsonify({'message': 'Pricing file updated successfully.'})

if __name__ == '__main__':
    print("="*55)
    print("  Peptide Take Home Generator")
    print("  Open: http://localhost:5000")
    print(f"  Staff login:    staff / {STAFF_USERS.get('staff','...')}")
    print(f"  Admin password: {ADMIN_PASSWORD}")
    print("  Ctrl+C to stop")
    print("="*55)
    app.run(host='0.0.0.0', port=5000, debug=False)
