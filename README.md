# Dairy Job Tracker MVP

A "One-Click" mobile-responsive web app built on Google Apps Script for tracking dairy job shifts.

## Overview
This application allows you to:
- Start and End shifts with a single button.
- Automatically calculate total hours.
- Sync data to a Google Sheet (Single Source of Truth).
- View recent shift history.
- Handle UTC+07:00 timezone correctly.

## Setup Instructions

1.  **Create a Google Sheet**:
    - Create a new Google Sheet.
    - **IMPORTANT: Rename the first tab to `ShiftData`.** (Default is usually 'Sheet1')
    - Set the header row (Row 1) with: `Shift ID`, `Start Time`, `End Time`, `Total Hours`.

2.  **Open Apps Script**:
    - Go to `Extensions` > `Apps Script`.

3.  **Deploy Backend Code**:
    - Build/Copy the content of `Code.js` into the default `Code.gs` file in the script editor.
    - Save the project.

4.  **Deploy Frontend Code**:
    - Create a new HTML file named `DairyShift.html` in the script editor.
    - Copy the content of `DairyShift.html` into this file.
    - Save.

5.  **Deploy as Web App**:
    - Click `Deploy` > `New deployment`.
    - Select type: `Web app`.
    - Configuration:
        - **Execute as**: `Me` (your account).
        - **Who has access**: `Anyone` (or `Anyone with Google account` if you want to restrict it).
    - Click `Deploy`.
    - Copy the "Web App URL".

## Local Usage (Optional)
If you want to run the app from your local computer (e.g., to keep it open on your desktop):

1. **Update Code**: Make sure you have redeployed the latest `Code.js` which supports the API.
2. **Open Local File**: Double-click `DairyShift_Local.html` on your computer to open it in your browser.
3. **Configure**:
    - At the top of the page, paste your **Web App URL** (from step 5 above).
    - It will remember this URL for next time.
    - You can now start/stop shifts locally, and it will sync to the Sheet!

## Architecture
- **Frontend**: 
    - `DairyShift.html`: Hosted on Google Servers.
    - `DairyShift_Local.html`: Local file using Fetch API.
- **Backend**: Google Apps Script (`.gs`).
- **Database**: Google Sheet.
