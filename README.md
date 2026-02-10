# Intelligent Job Board Scraper 📊

A professional, high-end job scraping dashboard and analytics suite designed for recruitment optimization. This application provides a seamless interface to submit scraping requests to an n8n-powered backend and visualizes the results with real-time analytics.

## 🚀 Features

- **Advanced Scraping Requests**: Custom form to target specific job titles, platforms (LinkedIn, Indeed, Glassdoor), locations, and salary ranges.
- **Live Analytics Dashboard**:
  - **Metric Cards**: Total jobs scraped, unique companies, top job types, and more.
  - **Interactive Charts**: Visual breakdown of jobs by company, city, match score distribution, and role category using Chart.js.
- **Dynamic Data Table**: 
  - Real-time data sync from Google Sheets.
  - Expandable cells for long job descriptions and match analysis.
  - **In-Dashboard Outreach**: Directly send emails via n8n webhooks from the data table.
- **Enterprise UI**: Clean, professional, emoji-free interface with a focus on usability and clarity.
- **Auto-Refresh**: Automatically stays in sync with your data source every 30 seconds.

## 🛠️ Tech Stack

- **Frontend**: HTML5, Vanilla CSS3 (Custom Variables, Modern Layouts), JavaScript (ES6+).
- **Visualization**: [Chart.js](https://www.chartjs.org/) for high-performance data visualization.
- **Backend Integration**: n8n Webhooks for data processing and email outreach.
- **Data Source**: Google Sheets (via CSV export API).

## 📋 Setup & Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/YOUR_USERNAME/intelligent-job-board-scraper.git
   cd intelligent-job-board-scraper
   ```

2. **Configuration**:
   Open `app.js` and update the `CONFIG` object with your production URLs:
   ```javascript
   const CONFIG = {
       webhookUrl: 'YOUR_N8N_WEBHOOK_URL',
       googleSheetUrl: 'YOUR_GOOGLE_SHEET_CSV_URL',
       emailWebhookUrl: 'YOUR_OUTREACH_WEBHOOK_URL',
       refreshInterval: 30000 
   };
   ```

3. **Running Locally**:
   Simply open `index.html` in your browser, or serve it using a local server:
   ```bash
   # Using Python
   python3 -m http.server 8080
   ```

## 🔐 Security Note

Ensure your n8n webhook endpoints are configured to allow CORS requests from the domain where this dashboard is hosted.

## 📄 License

This project is licensed under the MIT License.
