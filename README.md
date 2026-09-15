# 📊 OmniConverter • Multi-Format Call Log Converter & Analytics Dashboard

[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live%20Demo-success?logo=github)](https://jd164.github.io/omniconverter/)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-Framework-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

An all-in-one tool and interactive web dashboard to convert WhatsApp and phone call history exports into **Excel (.xlsx)**, **CSV**, **JSON**, and **iCalendar (.ics)** calendar events.

> 🔒 **Privacy-First**: The web dashboard is fully client-side capable. When hosted on **GitHub Pages**, your call data is processed directly inside your browser—no files or personal data are ever uploaded to any external server.

---

## 🌐 Use Online on GitHub Pages (No Install Required)

You can use this tool directly in your browser without installing anything:

👉 **[Launch OmniConverter on GitHub Pages](https://jd164.github.io/omniconverter/)**

1. Open the link above.
2. Drag & drop your call log `.json` file (or click **Load Sample Data** to explore the demo).
3. View instant KPIs, duration charts, and interactive call records.
4. Click any format button (**Excel**, **CSV**, **Calendar .ics**, **JSON**, or **ZIP bundle**) to download immediately!

---

## ⚡ Key Features

- **📊 Multi-Format Conversion**:
  - **Excel (.xlsx)**: Formatted workbook with two sheets:
    - `Call Log`: Filtered records with styled headers, alternating row colors, auto-adjusted column widths, and duration formatting.
    - `Statistical Summary`: Aggregated volume metrics, total hours, call averages, and date span.
  - **CSV (.csv)**: RFC-4180 compliant with **UTF-8 BOM (`\uFEFF`)** and `;` delimiter so Microsoft Excel opens Portuguese and international accents cleanly.
  - **iCalendar (.ics)**: RFC-5545 standard calendar event file with exact UTC timestamps, duration, and customizable summary template (e.g. `Call with {contact}`).
  - **JSON (.json)**: Normalized, indented, and clean structured data.
  - **ZIP Bundle (.zip)**: Download all 4 formats packaged together with a single click.
- **📈 Real-Time KPIs & Analytics**:
  - Total Calls & Analyzed Period
  - Total Accumulated Duration (Hours, Minutes, Seconds)
  - Average Duration per Completed Call
  - Voice vs Video Proportion (with visual progress indicator)
  - Answer Rate (Answered vs Missed calls)
- **📉 Visual Charts (`Chart.js`)**:
  - **Timeline Analytics**: Daily & monthly conversation volume and hours. Toggle between Duration (Hours) and Call Count.
  - **Peak Calling Hours**: Hourly distribution (00:00 to 23:00) showing when calls occur most frequently.
- **🔍 Dynamic Filtering & Search**:
  - Instant text search (matches ID, date, time, duration, status, direction).
  - Date range filtering (From date / To date).
  - Quick filter pills (`All`, `Voice`, `Video`, `Answered`, `Missed`, `> 30 min`).
  - Selective export: Export only currently filtered calls or specific row selections.
- **📋 Interactive Data Table**:
  - Multi-column sorting (Date, Time, Duration).
  - Configurable pagination (15, 30, 50, 100, or All items per page).
  - Row details modal with full metadata inspection and one-click JSON copying.

---

## 💻 Running Locally

If you prefer to run the application on your computer with the local Python server:

### Option 1: Double-Click Batch File (Windows)
Double-click:
```bat
start_dashboard.bat
```
The local server will start and automatically open your default browser at `http://127.0.0.1:8000`.

### Option 2: Terminal / Command Line
```bash
# Start local web dashboard
python main.py --web
```

---

## 🖥️ Command Line Interface (CLI)

You can also convert call logs directly in your terminal without opening a browser:

```bash
# Convert to all 4 formats at once (Excel, CSV, JSON, ICS):
python main.py

# Convert to a specific format:
python main.py --format xlsx   # Excel spreadsheet only
python main.py --format csv    # CSV file only
python main.py --format json   # JSON file only
python main.py --format ics    # iCalendar (.ics) only

# Custom input, output, and calendar event title:
python main.py -i my_calls.json -o report -f all -t "Call with {contact}"
```

### CLI Parameters:
| Option | Short | Default | Description |
| :--- | :--- | :--- | :--- |
| `--input` | `-i` | `dados.json` | Path to source JSON file |
| `--output` | `-o` | `call_logs` | Base name for exported files |
| `--format` | `-f` | `all` | Target: `all`, `xlsx`, `csv`, `json`, `ics` |
| `--title` | `-t` | `Call with {contact}` | Calendar event summary template |
| `--web` | | `False` | Launch interactive web dashboard |
| `--port` | | `8000` | Local server port |

---

## 📁 Repository Structure

```
├── index.html            # Main web dashboard entry point (GitHub Pages root)
├── style.css             # Glassmorphism dark mode stylesheet
├── app.js                # Reactive UI, Chart.js engine & client-side converters
├── sample_data.json      # Anonymized sample dataset for demo and testing
├── dados.json            # Anonymized default dataset
├── libs/                 # Local vendor libraries for offline operation
│   ├── chart.umd.min.js  # Chart.js
│   ├── jszip.min.js      # JSZip for ZIP downloads
│   └── xlsx.full.min.js  # SheetJS for Excel generation
├── docs/                 # GitHub Pages docs folder backup
├── app.py                # Local FastAPI backend server
├── converter.py          # Python conversion engine (openpyxl, csv, json, ics)
├── main.py               # Python CLI entrypoint & web launcher
├── start_dashboard.bat   # Windows one-click launcher
└── README.md             # Project documentation
```

---

## 🔒 Sensitive Data & Security

This repository contains **strictly anonymized sample data**. All phone numbers, personal identities, contact names, and internal WhatsApp JIDs have been sanitized. When using the application online or locally, your personal files remain stored only on your client machine.

---

## 📄 License

This project is licensed under the MIT License. Feel free to use, modify, and distribute.
