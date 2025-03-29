# 🕒 Timesheet Comparator

A FastAPI-based utility to compare SAP and WAND timesheets using a mapping file. It processes uploaded Excel/CSV files, analyzes the data, and returns a styled Excel report with comparisons, summaries, and charts.

---

## 📦 Installation

1. **Clone the repository**
```bash
git clone https://github.com/satyamvatstyagi/timesheet-comparator.git
cd timesheet-comparator
```

2. **(Optional) Create and activate a virtual environment**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

---

## 🚀 Running the App Locally

Start the FastAPI server using Uvicorn:
```bash
uvicorn app.main:app --reload
```

Once running, visit:
- Swagger UI: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)
- Redoc: [http://127.0.0.1:8000/redoc](http://127.0.0.1:8000/redoc)

---

## 📤 API Endpoint

### `POST /compare-timesheets/`

#### 📥 Upload Form Fields:
- `sap_file` - SAP Excel file (`.xlsx`)
- `wand_file` - WAND Excel file (`.xls`)
- `mapping_file` - CSV file mapping SAP and WAND users

#### 📤 Response:
- Returns an Excel file (`comparison_report.xlsx`) with:
  - Sheet 1: Detailed comparison
  - Sheet 2: Summary of discrepancies
  - Sheet 3: Visual charts (bar/line)

You can test this easily using Swagger UI.

---

## 🧪 Example Script (Test via Python)
```python
import requests

url = "http://127.0.0.1:8000/compare-timesheets/"

with open("sap.xlsx", "rb") as sap_file, \
     open("wand.xls", "rb") as wand_file, \
     open("mapping.csv", "rb") as mapping_file:

    files = {
        "sap_file": ("sap.xlsx", sap_file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "wand_file": ("wand.xls", wand_file, "application/vnd.ms-excel"),
        "mapping_file": ("mapping.csv", mapping_file, "text/csv")
    }

    response = requests.post(url, files=files)

    if response.status_code == 200:
        with open("comparison_report.xlsx", "wb") as out_file:
            out_file.write(response.content)
        print("✅ Report saved successfully!")
    else:
        print(f"❌ Failed: {response.status_code} - {response.text}")
```

---

## 🗂 Project Structure
```
timesheet-comparator/
├── app/
│   ├── main.py
│   ├── services/
│   │   └── timesheet_comparator.py
│   └── utils/
│       ├── excel_formatter.py
│       ├── load_mapping_file.py
│       ├── load_sap_file.py
│       ├── load_wand_file.py
│       └── write_report_file.py
├── report/          # Generated Excel reports saved here
├── test/
│   └── test_timesheet_comparator.py
├── requirements.txt
└── README.md
```

---

## 🙌 Contributions
Pull requests and suggestions welcome!

## 📧 Contact
For questions or support, email: `satyamvats@gmail.com`
