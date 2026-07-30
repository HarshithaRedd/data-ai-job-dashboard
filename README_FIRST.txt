DATA, AI & ANALYST JOB DASHBOARD — START HERE

This folder is complete. Do not copy files one by one.

FIRST TIME ONLY
1. Double-click SETUP_ONCE.bat.
2. Wait until it says Setup complete.

UPDATE THE JOBS
1. Double-click UPDATE_JOBS.bat.
2. It checks all 54 companies and creates the JSON, CSV, and Excel files.

OPEN THE DASHBOARD
1. Double-click OPEN_DASHBOARD.bat.
2. Keep the black window open while using the dashboard.
3. The dashboard opens at http://localhost:5500

TEST ONE COMPANY
Open PowerShell in this folder and run:
  venv\Scripts\Activate.ps1
  python main.py --company "Motion Recruitment"

AUTOMATIC ONLINE UPDATES
The included GitHub Actions workflow checks every hour at minute 17.
After pushing this folder to a public GitHub repository, set Settings > Pages > Source to GitHub Actions.

IMPORTANT
The collector uses official public career pages, public ATS APIs when available, HTML/JSON-LD, sitemaps, and a Playwright browser fallback. Some companies can block automation, require a CAPTCHA, or change their site. The dashboard's Source Health section shows exactly which companies succeeded, returned zero, or failed. It does not invent missing jobs.
