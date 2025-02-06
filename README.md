# Timesheet Comparator

This FastAPI utility allows users to upload SAP, WAND, and mapping files, processes the data, compares timesheets, and returns an Excel report.

## Installation
```sh
pip install -r requirements.txt
```

## Running the Server
```sh
uvicorn main:app --reload
```

## API Endpoint
- **POST** `/compare-timesheets/`
  - Upload SAP, WAND, and Mapping files.
  - Returns an Excel file with the comparison report.
"""