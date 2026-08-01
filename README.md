# Personal Dashboard

A modern, responsive, single-page dark-mode dashboard built with **Python Flask** and **Vanilla CSS/JS**. It consolidates your local environment widgets (live clock, weather) with your primary workflow tools: **Monday.com Projects** and **Microsoft Outlook Calendar**.

## Features

1. **Time & Greeting Widget**: Displays a real-time digital clock, localized date, and a welcoming header tailored to the hour of the day.
2. **Zero-Config Weather Widget**: Leverages a completely free, registration-less JSON service (`wttr.in`) to display temperature, conditions, humidity, and wind speed. You can easily toggle cities from the UI dropdown list.
3. **Monday.com Integration (Board 9609739665)**:
   - Queries and displays the 30 most recent active projects.
   - Shows **Sales Owner**, **Flag**, **Due Date**, and **Status** with custom designed color-coded badges matching Monday's native aesthetics.
4. **Outlook Calendar Integration**:
   - Safely reads your MSAL node credentials token cache from `outlook-mcp`.
   - If the access token has expired, it automatically requests a silent token refresh in the background.
   - Lists your upcoming meetings for the next 7 days, including times, locations, and organizers.
5. **Auto-Discovery Credentials**: Designed to run immediately with **zero-configuration**. It automatically discovers:
   - Microsoft App registration credentials (`CLIENT_ID` / `TENANT_ID`) by scanning your local `outlook-mcp/.env` file.
   - Monday.com API Token by parsing your local `monitor_monday.sh` script.

---

## Local Setup

### 1. Requirements
Ensure you have Python 3.9+ installed.

### 2. Install Dependencies
Navigate to the directory and install python requirements:
```bash
cd personal-dashboard
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Run the Dashboard
Start the development server:
```bash
python3 app.py
```
The application will run on **`http://localhost:5001`** (port 5001 is used to prevent any conflicts with default systems or local webapps).

---

## Git & Deployment to GitHub / Render

### Push to your GitHub Repository
Initialize git, connect to your repository, and push:
```bash
git init
git add .
git commit -m "Initial commit of Personal Dashboard"
git branch -M main
git remote add origin https://github.com/mattwegesin/personal-dashboard.git
git push -u origin main
```

### Hosting on Render
1. Log in to your [Render Dashboard](https://dashboard.render.com).
2. Click **New +** and select **Web Service**.
3. Connect your **GitHub account** and select the `personal-dashboard` repository.
4. Configure the settings:
   - **Environment**: `Python`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. Click **Deploy Web Service**.
6. Under **Environment**, add your environment variables:
   - `FLASK_SECRET_KEY`: A random string.
   - `MONDAY_API_TOKEN`: Your Monday.com token.
   - `BOARD_ID`: `9609739665`
   - `CLIENT_ID`: Your Microsoft Client ID.
   - `TENANT_ID`: Your Microsoft Tenant ID (or `common`).
