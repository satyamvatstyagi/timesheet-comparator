import requests

url = "http://127.0.0.1:8080/compare-timesheets/"

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
        with open("comparison_report_1.xlsx", "wb") as out_file:
            out_file.write(response.content)
        print("✅ Report saved as comparison_report_1.xlsx")
    else:
        print(f"❌ Failed: {response.status_code} - {response.text}")
