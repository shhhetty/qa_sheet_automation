# Fully Automated Google Sheet Validation Service

> An automated, cloud-deployed validation service that continuously monitors and compares product data between two Google Sheets, running 24/7 on a completely free-cost infrastructure.

## 1. The Business Problem
Our team faced a significant operational challenge: the need to constantly validate product assortment data between a master sheet and a partner's sheet. The manual process was not only slow (taking hours) and prone to human error, but it also wasn't continuous. Discrepancies could exist for hours or days before the next manual check, risking inventory and listing errors on live platforms. We needed a solution that was **always on, completely automated, and required zero manual intervention.**

## 2. My Solution & System Architecture
I engineered a complete, end-to-end automated service. This is not just a Python script, but a deployed application designed for reliability and cost-efficiency.

**A. Core Logic (Python):**
The heart of the service is a Python script that uses:
- **Gspread & OAuth2:** To securely authenticate and interact with the Google Sheets API.
- **Pandas:** To efficiently load data from both sheets into DataFrames for high-speed, in-memory comparison.
- **Discrepancy Reporting:** The logic identifies any rows with mismatched data based on a unique product key and writes a detailed report to a separate "Validation Report" sheet, complete with timestamps.

**B. Deployment as a Web Service (Render):**
To achieve continuous operation, I containerized the Python application and deployed it as a web service on **Render**.
- **Why Render?** Render's free tier for web services provides enough resources to run this application without any cost. It also offers auto-deploys from GitHub, meaning any updates I push to the main branch are automatically deployed to the live service.
- **Web Framework:** I used a lightweight framework like **Flask** to wrap the core logic in an API endpoint. This allows the validation task to be triggered by a simple HTTP request.

**C. Ensuring 24/7 Uptime (The "Keep-Alive" Strategy):**
Render's free web services are designed to "spin down" after 15 minutes of inactivity to save resources. To prevent this and ensure my validation service was always active, I implemented a clever keep-alive strategy:
- **Uptime Robot:** I configured a free monitor from **Uptime Robot** to send an HTTP GET request to my service's public URL every 5-10 minutes.
- **The Result:** This constant "ping" tricks the Render service into thinking it's always active, preventing it from spinning down. This guarantees the service is always ready to run its scheduled task or be triggered.

This multi-part architecture transforms a simple script into a robust, serverless automation engine.

## 3. Technologies & Skills
- **Primary Language:** Python
- **Libraries:** Gspread, OAuth2Client, Pandas, Flask
- **Cloud & DevOps:**
  - **Deployment:** Render
  - **CI/CD:** Git, GitHub (for auto-deploys)
  - **Monitoring/Uptime:** Uptime Robot
- **APIs:** Google Sheets API, REST API (for the service endpoint)
- **Core Skills:** QA Automation, Process Automation, Cloud Deployment, Systems Architecture, Resource Management

## 4. The Impact
- **Zero-Cost Operation:** The entire solution runs 24/7 with absolutely no infrastructure cost.
- **100% Automation:** The process requires zero human intervention after the initial setup.
- **Real-Time Alerts:** By running continuously, the service can detect and report discrepancies within minutes of their creation, not hours or days later.
- **Enhanced Reliability:** The deployed service is far more reliable than a script running on a local machine, which could be turned off or lose connection.

## 5. Replicating the Setup
1.  **Code:** Clone this repository.
2.  **Deployment:** Create a new "Web Service" on Render and connect it to your GitHub repo.
3.  **Build & Run:** Set the build command to `pip install -r requirements.txt` and the start command to `gunicorn app:app` (or similar, depending on your Flask setup).
4.  **Environment Variables:** Add your Google API credentials as secret files or environment variables in Render.
5.  **Uptime:** Create a free HTTP monitor in Uptime Robot pointing to your public Render URL (`https://your-service-name.onrender.com`).
