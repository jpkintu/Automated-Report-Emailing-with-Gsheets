# Automated Report Emailing with Google Sheets

This project automates the process of generating and emailing daily sales summary reports using Google Apps Script and Google Sheets.

## Features

- Extracts sales and OTP verification data from specified Google Sheets.
- Generates a styled HTML summary report.
- Identifies and highlights agents with low sales performance.
- Emails the report to designated recipients.
- Daily scheduling via time-based triggers.

## How It Works

1. **Data Gathering:**  
   Fetches data from "Bi-weekly Stats" and "SF Looker Update" sheets.

2. **Summary Generation:**  
   Computes daily sales totals and OTP rates for each product (ICS, EPC, WF).

3. **Performance Monitoring:**  
   Identifies agents with less than 20 ICS sales for targeted follow-up.

4. **Email Delivery:**  
   Sends a summary email with tables and highlights to the specified address.

5. **Automation:**  
   Uses Apps Script's `createTrigger()` to schedule automatic daily reports.

## Setup

1. **Clone or Copy Script:**  
   Add `dailyDistributions.gs` to your Google Apps Script project linked to your Google Sheet.

2. **Configure Sheet Names:**  
   Ensure your sheet names match `"Bi-weekly Stats"` and `"SF Looker Update"`.

3. **Set Recipient Email:**  
   Update the `to:` field in `MailApp.sendEmail()` to your preferred recipient.

4. **Set Daily Trigger:**  
   Run `createTrigger()` once to schedule daily execution at 10 AM.

## Customization

- Adjust thresholds or sheet names as needed in the script.
- Modify HTML/CSS styles in the `summaryHtml` section for branding or readability.
