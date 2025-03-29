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
    # === Load mapping file ===
    mapping_df = pd.read_csv(mapping_file.file)
    mapping_df.columns = mapping_df.columns.str.strip()
    email_col = next(
        (col for col in mapping_df.columns if 'email' in col.lower()), None)
    if email_col is None:
        raise ValueError("❌ 'email' column not found in mapping.csv")
    print("✅ Mapping loaded")

    # === Load SAP ===
    sap_raw = pd.read_excel(sap_file.file, header=None, skiprows=19)
    date_row = sap_raw.iloc[0, 6:].tolist()
    dates = pd.to_datetime(date_row, errors='coerce').date
    sap_data = sap_raw.iloc[2:].reset_index(drop=True)

    emails = sap_data.iloc[:, 4]
    full_names = sap_data.iloc[:, 3]
    hours_data = sap_data.iloc[:, 6:]
    hours_data = hours_data.applymap(lambda x: float(
        str(x).replace(" H", "").strip()) if pd.notnull(x) else 0)

    # === Melt SAP
    sap_long = pd.melt(
        hours_data.reset_index(drop=True),
        var_name="DayIndex",
        value_name="Hours"
    )
    sap_long['Email'] = emails.repeat(
        hours_data.shape[1]).reset_index(drop=True)
    sap_long['Full Name'] = full_names.repeat(
        hours_data.shape[1]).reset_index(drop=True)
    sap_long['Date'] = list(dates) * len(sap_data)
    sap_long = sap_long[['Email', 'Full Name', 'Date', 'Hours']].dropna(subset=[
                                                                        'Date', 'Hours'])

    # === Group SAP and enrich with mapping
    sap_grouped = sap_long.groupby(['Email', 'Full Name', 'Date']).agg({
        'Hours': 'sum'}).reset_index()
    email_col = next(
        (col for col in mapping_df.columns if 'email' in col.lower()), None)
    if email_col is None:
        raise ValueError("❌ 'email' column not found in mapping.csv")
    sap_grouped = sap_grouped.merge(
        mapping_df, left_on='Email', right_on=email_col, how='left')
    sap_grouped.drop(columns=['Email'], inplace=True)

    # Rename detected email column to 'emailAddress' for consistency
    sap_grouped.rename(columns={email_col: 'emailAddress'}, inplace=True)

    # === Load WAND ===
    wand_df = pd.read_excel(wand_file.file, engine='xlrd')
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

    # === Enrich WAND with mapping
    wand_long = wand_long[wand_long['Name'].isin(mapping_df['proWandName'])]
    wand_long = wand_long.merge(
        mapping_df, left_on='Name', right_on='proWandName', how='left')

    wand_grouped = wand_long.groupby(['emailAddress', 'Date', 'projectName']).agg({
        'Hours': 'sum'}).reset_index()

    # === Merge SAP and WAND
    merged = pd.merge(
        sap_grouped,
        wand_grouped,
        on=['emailAddress', 'Date', 'projectName'],
        how='outer',
        suffixes=('_sap', '_wand')
    )

    merged['Hours_sap'] = merged['Hours_sap'].fillna(0)
    merged['Hours_wand'] = merged['Hours_wand'].fillna(0)
    merged['Delta'] = merged['Hours_sap'] - merged['Hours_wand']

    # === Clean column names
    merged.rename(columns={'emailAddress': 'Email'}, inplace=True)

    # ✅ Reorder columns
    desired_order = ['Date', 'Email', 'Full Name', 'proWandName',
                     'projectName', 'Hours_sap', 'Hours_wand', 'Delta']
    merged = merged[desired_order]

    # === Export
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged.to_excel(writer, index=False, sheet_name='Comparison Report')

    output.seek(0)
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": "attachment; filename=comparison_report.xlsx"}
    )
