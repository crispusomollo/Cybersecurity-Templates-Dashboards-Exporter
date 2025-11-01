# 🛡️ Cybersecurity Templates and Dashboards Exporter

This project automatically generates Cybersecurity documentation and dashboards across all security lifecycle phases.  
It includes ready-made `.docx`, `.xlsx`, and `.csv` templates — complete with sample data, charts, and an executive summary.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![License: MIT](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/status-stable-brightgreen)
![Last Commit](https://img.shields.io/github/last-commit/crispusomollo/cybersecurity-templates-dashboards-exporter)
![Open Issues](https://img.shields.io/github/issues/crispusomollo/cybersecurity-templates-dashboards-exporter)
![Repo Size](https://img.shields.io/github/repo-size/crispusomollo/cybersecurity-templates-dashboards-exporter)
![Top Language](https://img.shields.io/github/languages/top/crispusomollo/cybersecurity-templates-dashboards-exporter)

---

## 🚀 Features
- Generates 5 cybersecurity phase folders automatically
- Creates:
  - Documentation templates (`.docx`)
  - Data dashboards (`.xlsx`) with charts
  - Export logs (`.csv`)
  - Executive Summary Report
- Produces a ready-to-share `.zip` archive
- Works entirely offline

---

## 📦 Installation

Clone the repo and install dependencies:

```bash
git clone https://github.com/crispusomollo/cybersecurity-templates-dashboards-exporter.git
cd cybersecurity-templates-dashboards-exporter
pip install -r requirements.txt
```

---

## ▶️ Usage

Run the generator script:

```
python generate_cybersecurity_templates.py
```

This will create a folder and ZIP archive:

```
Cybersecurity_Templates_&_Dashboards_Exporter/
Cybersecurity_Templates_&_Dashboards_Exporter.zip
```

---

## 📊 Output Example

Each phase contains:

```
Phase_1_Governance_Risk_SOC/
 ├── Phase_1..._Template.docx
 ├── Phase_1..._Dashboard.xlsx  (with charts)
 └── Phase_1..._Data.csv
```

The root folder includes:

```
Cybersecurity_Executive_Summary.docx
Supporting_Files/export_summary.csv
```

---

## 🧩 Requirements

- Python 3.9+
- pandas
- openpyxl
- python-docx
- tqdm

---

## 📜 License

MIT License © 2025 Crispus Omollo
Feel free to use, modify, and distribute.

