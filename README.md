# ⚡ CGNET Employee Audit Scraper & Web Portal

A modern, mobile-responsive Web Application and automated scraper built with Python, Streamlit, and Playwright for extracting CGNET employee performance audit records, calculating SLA resolution metrics, and generating Excel reports.

---

## 📱 How to Access & Host for FREE from your Mobile Phone

### Option 1: 100% Free Cloud Hosting (Streamlit Community Cloud) — *Recommended*

You can host this app online for free in less than 3 minutes, giving you a custom web link (e.g. `https://your-name-audit.streamlit.app`) accessible from your iPhone, Android, or laptop anywhere!

1. **Upload your code to GitHub**:
   - Create a free account on [GitHub.com](https://github.com).
   - Create a new public or private repository (e.g., `employee-audit-scraper`).
   - Push these files to your repository:
     - `app.py`
     - `audit_automation.py`
     - `requirements.txt`
     - `packages.txt`

2. **Deploy on Streamlit Community Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io) and log in with GitHub.
   - Click **"New app"**.
   - Select your repository (`employee-audit-scraper`), branch (`main`), and set **Main file path** to `app.py`.
   - Click **Deploy!**

3. **Access from Phone**:
   - Open your generated app URL on your phone browser, choose any date range, and tap **Run Audit Scraper** to view and download reports!

---

### Option 2: Instant Phone Access via Free Tunnel (Local PC Run)

If you prefer running the browser on your computer while accessing the UI from your phone:

1. **Start the Web App on your PC**:
   ```bash
   uv run streamlit run app.py
   ```

2. **Expose to Phone via Cloudflare Tunnel (Free)**:
   In a separate terminal on your PC, run:
   ```bash
   npx localtunnel --port 8501
   ```
   *or*
   ```bash
   npx @cloudflare/cloudflared tunnel --url http://localhost:8501
   ```

3. Open the provided `https://...` link on your phone!

---

## 🚀 Local Desktop Usage

To run the web interface directly on your local computer:

```bash
uv run streamlit run app.py
```

Or run via CLI:

```bash
uv run audit_automation.py --employee "Om Neupane" --from-date "04 Aug 2026" --to-date "05 Aug 2026" --non-interactive
```
