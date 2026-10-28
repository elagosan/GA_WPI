# WORK PLACE INTEGRATION GENERAL ASSEMBELY
### Standard Criteria & Station Assessment Dashboard

Interactive Streamlit executive analytics dashboard for Stellantis General Assembly WPI (Workplace Integration) implementation, Station Assessments, and Standard Criteria Tracking.

## 🚀 Live Site Deployment (Streamlit Community Cloud)
This dashboard is ready to deploy live for free on **Streamlit Community Cloud**:
1. Log in to [share.streamlit.io](https://share.streamlit.io/) with your GitHub account.
2. Click **Create app** > **Yup, I have an app**.
3. Select:
   - **Repository:** `elagosan/GA_WPI`
   - **Branch:** `main`
   - **Main file path:** `main.py`
   - **App URL:** Customize if desired
4. Click **Deploy!**

Your dashboard will be live on a public URL (e.g. `https://ga-wpi.streamlit.app`).

## 📊 Live Dynamic Chart Updates
- When any user uploads a new Excel file directly on the live site, the dashboard dynamically reads the data and **instantly recalculates all charts**.
- Features 50% / 50% synchronized side-by-side view for Station Assessment and Overall Criteria with 2D clustered charts and complete breakdown tables.
- Optional GitHub Auto-Sync: Add your GitHub Personal Access Token to Streamlit Cloud **Settings > Secrets** as `GITHUB_TOKEN = "your_token"` to automatically commit any newly uploaded Excel file back to GitHub.

## 💻 Local Run
```bash
pip install -r requirements.txt
streamlit run main.py
```
