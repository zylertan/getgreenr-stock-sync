# Deploy the GetGreenr Stock Tool as a website (free)

Goal: a link like `https://mm-getgreenr-stock.streamlit.app` your team opens in any
browser — no Terminal, no install. Hosting is free via **Streamlit Community Cloud**.
One-time setup ~15 minutes.

## Files this website needs

Put these five files together (nothing else is required):

- `getgreenr_app.py`  ← the website (main file)
- `getgreenr_core.py`
- `matcher.py`
- `registry_report.py`
- `requirements.txt`

---

## Step 1 — Create a free GitHub account (holds the code)

1. Go to https://github.com and click **Sign up**. Use your work email.
2. Verify your email.

## Step 2 — Put the files on GitHub (no coding)

1. Click the **+** (top-right) → **New repository**.
2. Repository name: `mm-getgreenr-stock`. Set it to **Private**. Click **Create repository**.
3. On the next page click **uploading an existing file** (the link in the middle).
4. Drag in the five files listed above. Click **Commit changes**.

## Step 3 — Deploy on Streamlit Community Cloud

1. Go to https://share.streamlit.io and click **Sign in** → **Continue with GitHub** (approve access).
2. Click **Create app** → **Deploy a public app from a repo**.
3. Fill in:
   - **Repository:** `your-username/mm-getgreenr-stock`
   - **Branch:** `main`
   - **Main file path:** `getgreenr_app.py`
4. Before deploying, click **Advanced settings** → **Secrets**, and paste this one line
   (choose your own password):

   ```
   APP_PASSWORD = "pick-a-team-password"
   ```

5. Click **Deploy**. Wait ~2 minutes for it to build. You'll get your URL.

## Step 4 — Share with your teammates

- Send them the URL + the team password.
- To restrict who can even open the page (recommended), in the app's **Settings →
  Sharing**, set viewers to specific emails (your teammates' Google/GitHub emails).
  That plus the password gives you two layers.

---

## Using it

Open the URL → enter the team password → upload the four files
(Masterlist, GetGreenr bulk stock, SKU Registry, and optionally New Masterlist SKUs)
→ click **Apply rules to Seller Stock** → download the updated GetGreenr file.

## Updating the app later

Edit a file on GitHub (or re-upload it) and commit — Streamlit redeploys automatically
within a minute. To change the password, edit the secret in **Settings → Secrets**.

## Notes

- Files uploaded in the app are processed in memory for that session and are **not**
  stored on the server.
- The password is only read from Streamlit **Secrets** — it is never in the code on GitHub.
- Free tier is fine for a small team; the app may "sleep" after inactivity and take ~30s
  to wake on the next visit.
