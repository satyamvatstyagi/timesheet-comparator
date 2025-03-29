import os
import requests


def test_compare_timesheets():
    url = "http://127.0.0.1:8000/compare-timesheets/"
    files = {
        "sap_file": open("sap.xlsx", "rb"),
        "wand_file": open("wand.xls", "rb"),
        "mapping_file": open("mapping.csv", "rb")
    }
    response = requests.post(url, files=files)

    assert response.status_code == 200
    with open("comparison_report_test.xlsx", "wb") as f:
        f.write(response.content)
    assert os.path.exists("comparison_report_test.xlsx")
