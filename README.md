# CBS Portfolio Dashboard Builder
## Deploying to Streamlit Community Cloud

This is the web app that lets anyone on your team upload the four source Excel files
and download a fresh HTML dashboard — no Python, no command line required.

---

### What you need (one-time setup, ~10 minutes)

- A free GitHub account — [github.com](https://github.com)
- A free Streamlit Cloud account — [share.streamlit.io](https://share.streamlit.io)

---

### Step 1 — Create a GitHub repository

1. Go to [github.com/new](https://github.com/new)
2. Give the repository any name, e.g. `cbs-dashboard-builder`
3. Set it to **Private** (recommended)
4. Click **Create repository**

---

### Step 2 — Upload these files to the repository

Upload all three files from this folder:

| File | Purpose |
|---|---|
| `app.py` | The web app |
| `build_core.py` | The dashboard build engine |
| `requirements.txt` | Python packages |

To upload: open your new repository on GitHub → click **Add file** → **Upload files** → drag all three files → click **Commit changes**.

---

### Step 3 — Connect Streamlit Cloud to GitHub

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
2. Click **New app**
3. Select your repository (`cbs-dashboard-builder`)
4. Set **Main file path** to `app.py`
5. Click **Deploy**

Streamlit will build and launch the app automatically (takes about 2 minutes the first time).

---

### Step 4 — Share the URL with your team

Once deployed, you get a permanent URL like:

```
https://your-name-cbs-dashboard-builder-app-xxxxxx.streamlit.app
```

Share this link with anyone on your team. They can open it in any browser with no login.

---

### Step 5 — Using the app

1. Open the URL in any browser
2. Upload the four Excel files (drag and drop or click to browse)
3. Click **Build Dashboard**
4. Click **Download Dashboard** to save the HTML file
5. Open or share the downloaded HTML file — it works in any browser, no internet required

---

### Updating the app after code changes

If `build_core.py` or `app.py` ever needs to be updated:

1. Upload the new file to the same GitHub repository (replace the old one)
2. Streamlit Cloud detects the change and redeploys automatically within ~1 minute

---

### Troubleshooting

| Symptom | Fix |
|---|---|
| "Build failed" error | Check that all four files are the correct latest exports |
| App is slow to load the first time | Normal — Streamlit Cloud "wakes up" cold apps after inactivity (~30 s) |
| File uploader does not accept the file | Make sure the files are `.xlsx` format, not `.csv` or `.xls` |
| Downloaded HTML is blank | Re-upload the source files and rebuild — the previous session may have expired |
