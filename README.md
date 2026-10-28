# WORKPLACE INTEGRATION – GENERAL ASSEMBLY
### Standard Criteria & Station Assessment Executive Dashboard

Interactive Streamlit executive analytics dashboard for Stellantis General Assembly WPI (Workplace Integration) implementation, Station Assessments, and Standard Criteria Tracking.

---

## 🏗️ Dashboard Architecture & Visual Flow
1. **Executive Header**: Title, General Assembly identification, Stellantis corporate branding, and live GitHub synchronization status chip.
2. **Station Assessment (Top Section)**:
   - Placed directly below the title header for immediate visibility.
   - 4 Station KPI Cards: **OVERALL STATIONS**, **IN PROGRESS**, **COMPLETED**, and **COMPLETION RATE %**.
   - Balanced side-by-side view: Interactive Milestone Progress Bar, Expandable Station Details Table, and Station Assessment Status Bar Chart.
   - Interactive Station Assessment metrics editor with GitHub commit sync.
3. **Dashboard Summary & Scope Counts**:
   - Observations scope metrics: Affected Observations, Closed Items, In Progress Items, Criteria Categories.
4. **Standard Criteria Overview**:
   - Full-width interactive 2D Bar Chart across all four standard criteria (Customer Experience, Workplace Organization, Health and Safety, Logistics).
   - Summary percentage pills and distribution breakdown.
5. **Standard Criteria Detailed Analytics**:
   - Paired charts for Health and Safety, Customer Experience, Workplace Organization, and Logistics.
   - Left: Status Bar Chart (Closed, In Progress, Affected).
   - Right: Action Category Breakdown (Standard, Facility, Process, Part, Training).
6. **Raw Data Explorer & Export**:
   - Filtered observation data inspection table.
   - Instant export to CSV and formatted multi-sheet Excel (.xlsx).

---

## 🚀 Live Site Deployment (Streamlit Community Cloud)
This dashboard is ready to deploy live for free on **Streamlit Community Cloud**:
1. Log in to [share.streamlit.io](https://share.streamlit.io/) with your GitHub account.
2. Click **Create app** > **Yup, I have an app**.
3. Select:
   - **Repository:** `elagosan/GA_WPI`
   - **Branch:** `main`
   - **Main file path:** `main.py`
4. In Streamlit Cloud **App Settings > Secrets**, add your GitHub Personal Access Token:
   ```toml
   GITHUB_TOKEN = "ghp_yourPersonalAccessTokenHere"
   ```
5. Click **Deploy!**

---

## ⚡ Instant GitHub Auto-Sync & Real-Time Notifications
When any user uploads an Excel workbook directly on the site:
- **🐙 Loaded in GitHub**: The workbook is automatically committed and pushed to `elagosan/GA_WPI` on branch `main` via the GitHub API or Git CLI.
- **⚙️ Processing Data**: Automatically scans all sheets, detects headers even if offset by title banners, and standardizes criteria and categories.
- **📊 Retrieving Data**: Filters, normalizes, and prepares observation records and station assessments with high-speed in-memory caching.
- **📈 Generating Charts**: Instantly generates all 2D charts and automatically reloads the site to display updated results.
- **🔔 Live Notifications**: Real-time progress updates via animated status containers and toast notifications guide the user through every step.

---

## 💻 Local Run
```bash
# Clone the repository
git clone https://github.com/elagosan/GA_WPI.git
cd GA_WPI

# Install required dependencies
pip install -r requirements.txt

# Run the Streamlit application
streamlit run main.py
```
