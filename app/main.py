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
    # Step 1: Read SAP (no headers)
    sap_df_raw = pd.read_excel(sap_file.file, header=None)
    print("SAP Raw Shape:", sap_df_raw.shape)
    print("SAP Raw Data:", sap_df_raw.head())
    print("SAP Columns:", sap_df_raw.columns.tolist())
    print("SAP Data Types:", sap_df_raw.dtypes)

    # Step 2: Transpose
    sap_df_t = sap_df_raw.T
    print("Transposed SAP Shape:", sap_df_t.shape)
    print("Transposed SAP Data:", sap_df_t.head())
    print("Transposed SAP Columns:", sap_df_t.columns.tolist())
    print("Transposed SAP Data Types:", sap_df_t.dtypes)

    # Step 3: Set headers from row 0
    sap_df_t.columns = sap_df_t.iloc[0]     # First row becomes header
    sap_df_t = sap_df_t[1:]                 # Drop header row

    # Step 4: Melt into long format
    sap_long = pd.melt(
        sap_df_t.reset_index(),
        id_vars=['index'],
        var_name='Name',
        value_name='Hours'
    )

    # Rename 'index' to 'Date'
    sap_long = sap_long.rename(columns={'index': 'Date'})

    # Step 5: Clean and format
    sap_long['Date'] = pd.to_datetime(
        sap_long['Date'], errors='coerce').dt.date
    sap_long = sap_long.dropna(subset=['Hours'])

    # --- Load and melt WAND Timesheet ---
    wand_df_raw = pd.read_excel(wand_file.file)
    print("WAND Columns:", wand_df_raw.columns.tolist())
    print("WAND Raw Shape:", wand_df_raw.shape)
    print("WAND Raw Data:", wand_df_raw.head())
    print("WAND Data Types:", wand_df_raw.dtypes)

    wand_long = pd.melt(
        wand_df_raw,
        id_vars=['Name'],
        var_name='Date',
        value_name='Hours'
    )
    wand_long = wand_long.dropna(subset=['Hours'])
    wand_long['Date'] = pd.to_datetime(wand_long['Date']).dt.date

    # --- Load Mapping File ---
    mapping_df = pd.read_csv(mapping_file.file)
    print("Mapping Columns:", mapping_df.columns.tolist())

    # Filter WAND based on mapping
    wand_long = wand_long[wand_long['Name'].isin(mapping_df['proWandName'])]
    wand_long = wand_long.merge(
        mapping_df, left_on='Name', right_on='proWandName', how='left')

    # Convert Hours to numeric to avoid type issues
    sap_long['Hours'] = pd.to_numeric(sap_long['Hours'], errors='coerce')
    sap_long = sap_long.dropna(subset=['Hours'])

    wand_long['Hours'] = pd.to_numeric(wand_long['Hours'], errors='coerce')
    wand_long = wand_long.dropna(subset=['Hours'])

    # --- Group and sum hours ---
    sap_grouped = sap_long.groupby(['Name', 'Date']).agg(
        {'Hours': 'sum'}).reset_index()
    wand_grouped = wand_long.groupby(
        ['Name', 'Date', 'emailAddress', 'projectName']).agg({'Hours': 'sum'}).reset_index()

    # --- Merge and calculate delta ---
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

    # --- Save to Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged.to_excel(writer, index=False, sheet_name='Comparison Report')

    output.seek(0)

    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': 'attachment; filename=comparison_report.xlsx'}
    )
