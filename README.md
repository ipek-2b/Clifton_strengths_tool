# CliftonStrengths Pipeline — Web App

Drag-and-drop tool: upload CliftonStrengths files (PDF/Excel/CSV/pasted text), get back a data spreadsheet and a wheel-chart spreadsheet. No installation for end users — just a link in a browser.

## Files in this folder

- `app.py` — the web app (Streamlit)
- `clifton_strengths_pipeline.py` — the parsing/spreadsheet/chart engine (same one used from the command line)
- `requirements.txt` — dependencies for hosting

## Deploy it (free, ~10 minutes, one-time setup)

**1. Put this folder on GitHub**

If you don't have a GitHub account, make one free at github.com/join.

- Go to github.com → **New repository** → name it e.g. `clifton-strengths-tool` → **Create repository**
- On the new repo's page, click **uploading an existing file**, drag in all three files from this folder (`app.py`, `clifton_strengths_pipeline.py`, `requirements.txt`), then **Commit changes**

(If you're comfortable with git instead of the browser upload: `git init && git add . && git commit -m "init" && git remote add origin <your-repo-url> && git push -u origin main`)

**2. Deploy on Streamlit Community Cloud (free hosting)**

- Go to **share.streamlit.io** → sign in with your GitHub account
- Click **New app**
- Pick your repo, branch `main`, and set **Main file path** to `app.py`
- Click **Deploy**

First build takes 2-3 minutes. You'll get a public link like `https://your-app-name.streamlit.app` — that's what you share with the team. Anyone with the link can open it, drop in files, and download results. No login needed for them.

**3. Updating it later**

Push any change to the GitHub repo (edit `clifton_strengths_pipeline.py`, upload a new file, whatever) and Streamlit Cloud redeploys automatically within a minute or two.

## Optional: restrict who can access it

By default the link is public (anyone with it can use the tool, but no one can see other people's uploads or results — each visit is a private session). If you'd rather only your team can open it: in the app's Streamlit Cloud settings → **Sharing**, switch from "Public" to "Only specific people" and add their emails.

## Testing it yourself first (optional)

If you want to try it on your own computer before deploying:
```
pip install -r requirements.txt
streamlit run app.py
```
This opens the app in your browser at `localhost:8501`.

## Note

I built and tested the underlying parsing/spreadsheet/chart engine directly (it's the same code as the command-line version, already verified against real and synthetic data). I could not run the Streamlit interface itself inside this sandbox, since it has no internet access to install the `streamlit` package — it will run normally once deployed, since Streamlit Cloud installs `requirements.txt` on its own servers. Worth a quick test upload once it's live to confirm everything looks right end-to-end.
