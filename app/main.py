from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
import pandas as pd
from io import BytesIO

app = FastAPI()


@app.post("/compare-timesheets/")
async def compare_timesheets(
    sap_file: UploadFile = File(...),
    wand_file: UploadFile = File(...),
    mapping_file: UploadFile = File(...)
):
    # === Load mapping.csv ===
    mapping_df = pd.read_csv(mapping_file.file)
    print("✅ Mapping loaded:", mapping_df.columns.tolist())

    # === Load SAP timesheet (skip first 5 metadata rows) ===
    sap_df = pd.read_excel(sap_file.file, skiprows=5)
    sap_df.rename(columns={sap_df.columns[0]: 'Date'}, inplace=True)

    # Melt SAP to long format: Date | Name | Hours
    sap_long = pd.melt(
        sap_df,
        id_vars=['Date'],
        var_name='Name',
        value_name='Hours'
    )
    sap_long['Date'] = pd.to_datetime(
        sap_long['Date'], errors='coerce').dt.date
    sap_long['Hours'] = pd.to_numeric(sap_long['Hours'], errors='coerce')
    sap_long = sap_long.dropna(subset=['Date', 'Hours'])
    print("✅ SAP entries:", sap_long.shape)

    # === Load and melt WAND ===
    wand_df = pd.read_excel(wand_file.file, engine='xlrd')  # for .xls support
    wand_long = pd.melt(
        wand_df,
        id_vars=['Name'],
        var_name='Date',
        value_name='Hours'
    )
    wand_long['Date'] = pd.to_datetime(
        wand_long['Date'], errors='coerce').dt.date
    wand_long['Hours'] = pd.to_numeric(wand_long['Hours'], errors='coerce')
    wand_long = wand_long.dropna(subset=['Date', 'Hours'])
    print("✅ WAND entries:", wand_long.shape)

    # === Filter and enrich WAND using mapping ===
    wand_long = wand_long[wand_long['Name'].isin(mapping_df['proWandName'])]
    wand_long = wand_long.merge(
        mapping_df, left_on='Name', right_on='proWandName', how='left')

    # === Group both timesheets ===
    sap_grouped = sap_long.groupby(['Name', 'Date']).agg(
        {'Hours': 'sum'}).reset_index()
    wand_grouped = wand_long.groupby(
        ['Name', 'Date', 'emailAddress', 'projectName']
    ).agg({'Hours': 'sum'}).reset_index()

    # === Merge and compute delta ===
    merged = pd.merge(
        sap_grouped,
        wand_grouped,
        on=['Name', 'Date'],
        how='outer',
        suffixes=('_sap', '_wand')
    )
    merged['Hours_sap'] = merged['Hours_sap'].fillna(0)
    merged['Hours_wand'] = merged['Hours_wand'].fillna(0)
    merged['Delta'] = merged['Hours_sap'] - merged['Hours_wand']

    # === Export to Excel ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged.to_excel(writer, index=False, sheet_name='Comparison Report')

    output.seek(0)
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": "attachment; filename=comparison_report.xlsx"}
    )
