import os
import requests


def test_compare_timesheets():
    url = "http://127.0.0.1:8080/compare-timesheets/"
    files = {
        "sap_file": open("sap.xlsx", "rb"),
        "wand_file": open("wand.xls", "rb"),
        "mapping_file": open("mapping.csv", "rb")
    }
    response = requests.post(url, files=files)

    assert response.status_code == 200
    report_path = os.path.join("reports", "comparison_report.xlsx")
    with open(report_path, "wb") as f:
        f.write(response.content)
    assert os.path.exists("comparison_report_test.xlsx")


if __name__ == "__main__":
    test_compare_timesheets()
    print("Test completed successfully.")
