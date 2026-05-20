# Deploying to Railway (Recommended — Free Tier)

Railway is the simplest cloud host for this app. Anyone with a browser
can use the tool once it's deployed — no Python needed on their computer.

---

## What you need first

- A free account at https://railway.app (sign up with Google or GitHub)
- All 6 files in one folder on your computer:
  - app.py
  - peptide_generator.py
  - requirements.txt
  - Procfile
  - FINAL FINAL Pricing 100 per bottle.xlsx
  - Coding Take Home Instrucitons 3ml.docx
  - Tesamorelin Flyer.docx
  - BPC TB Flyer.docx

---

## Step-by-step deployment

### Step 1 — Create a GitHub repository
1. Go to https://github.com and sign in (or create a free account)
2. Click the green "New" button to create a repository
3. Name it something like `peptide-generator`, set it to **Private**
4. Click "Create repository"
5. Upload all 8 files from your folder using the "uploading an existing file" link

### Step 2 — Deploy on Railway
1. Go to https://railway.app and log in
2. Click **"New Project"**
3. Choose **"Deploy from GitHub repo"**
4. Select your `peptide-generator` repository
5. Railway will detect it as a Python app and start deploying automatically
6. Wait about 2 minutes for the build to finish

### Step 3 — Set your passwords (important!)
In Railway, click your project → **Settings** → **Variables**, then add:

| Variable       | Value                          |
|----------------|-------------------------------|
| `SECRET_KEY`   | Any long random string, e.g. `xK9#mP2$qL7nR4@vT` |
| `ADMIN_PASSWORD` | Your chosen admin password   |
| `STAFF_USERS`  | `{"front_desk":"yourpass","nurse":"yourpass2"}` |

> Add as many staff usernames/passwords as you need in STAFF_USERS.

### Step 4 — Get your URL
Railway gives you a public URL like `https://peptide-generator-production.up.railway.app`
Share this with your staff — they just open it in any browser and sign in.

---

## Updating the Excel file (no redeploy needed)

1. Open the app URL in your browser and sign in
2. Click **⚙ Admin** at the bottom of the page
3. Enter the admin password
4. Choose your new `.xlsx` file and click Upload
5. The old file is backed up automatically, the new one takes effect immediately

---

## Adding or changing staff passwords

In Railway → Settings → Variables, edit the `STAFF_USERS` value.
Format: `{"username1":"password1","username2":"password2"}`
Click **Save** — Railway restarts the app in about 30 seconds.

---

## Running locally (no internet needed)

```
python app.py
```
Then open http://localhost:5000 in your browser.
Default login: staff / peptide2026!

