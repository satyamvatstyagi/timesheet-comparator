import requests
import os

# URL of your FastAPI endpoint
url = "http://127.0.0.1:8080/compare-timesheets/"

# File paths
sap_path = "reports/sap.xlsx"
wand_path = "reports/wand.xls"
mapping_path = "reports/mapping.csv"

# Prepare files for upload
with open(sap_path, "rb") as sap_file, \
        open(wand_path, "rb") as wand_file, \
        open(mapping_path, "rb") as mapping_file:

    files = {
        "sap_file": ("sap.xlsx", sap_file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "wand_file": ("wand.xls", wand_file, "application/vnd.ms-excel"),
        "mapping_file": ("mapping.csv", mapping_file, "text/csv")
    }

    print("⏳ Uploading files and calling API...")

    response = requests.post(url, files=files)

    if response.status_code == 200:
        # Save path for physical file
        report_path = os.path.join("reports", "comparison_report.xlsx")
        with open(report_path, "wb") as out_file:
            out_file.write(response.content)
        print(f"✅ Report saved as {report_path}")
    else:
        print(f"❌ Request failed [{response.status_code}]")
        print(response.text)
