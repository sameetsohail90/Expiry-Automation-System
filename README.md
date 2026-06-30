# Device Expiry Auto-Updater

Django-based automation tool that bulk-updates "Expiry Date" fields across multiple device admin panels (even across different client sites/logins) by reading device edit-page links from an Excel file and driving the update through Selenium browser automation.

## Features

- Upload an Excel file (`.xlsx`) containing a list of device edit-page links — supports 50+ links in a single run.
- Supports multiple sites/clients in the same Excel file, each with its own Login URL, Username, and Password (optional columns — falls back to default credentials if left blank).
- Automatically logs into each site once, then processes all device links belonging to that site.
- Updates the expiry date/time field on each device's edit page to either:
  - A specific date/time chosen by the user (via the web form), or
  - Automatically calculated as "current date + 1 month".
- Handles JavaScript-based save buttons (`<a onclick="save();">`), readonly fields, and custom date pickers via direct DOM manipulation and JS function invocation.
- Outputs a result Excel file with a Status column (Success / Failed + reason) and the new expiry value for each processed link.
- Built-in page element discovery/debug mode that logs all form fields and buttons found on each page, to simplify configuring selectors for a new site.

## Tech Stack

- Python 3
- Django (web interface for file upload and triggering the automation)
- Selenium + webdriver-manager (browser automation)
- openpyxl (Excel read/write)

## How It Works

1. User uploads an Excel file via the Django web form (optionally selecting a manual expiry date/time).
2. The app reads the `Link` column (and optional `Login URL` / `Username` / `Password` columns) from the Excel file.
3. Selenium opens a Chrome browser, logs into each distinct site, navigates to each device's edit page, updates the expiry field, and saves.
4. A result Excel file is generated and returned for download, showing the outcome of each update.

## Setup

```bash
pip install django selenium openpyxl webdriver-manager python-dateutil
python manage.py migrate
python manage.py runserver
```

Then visit `http://127.0.0.1:8000/update-expiry/` and upload an Excel file with a `Link` column.

## Disclaimer

This tool is intended for internal/authorized automation of accounts and systems you own or have explicit permission to manage. Selector configuration (login fields, expiry field, save button) needs to be adjusted to match the target site's HTML structure.
