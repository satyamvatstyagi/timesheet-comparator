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
        wand_long['Date'], errors='coerce')
    wand_long['Hours'] = pd.to_numeric(wand_long['Hours'], errors='coerce')
    wand_long = wand_long.dropna(subset=['Date', 'Hours'])

    # === Enrich WAND with mapping
    wand_long = wand_long[wand_long['Name'].isin(mapping_df['proWandName'])]
    wand_long = wand_long.merge(
        mapping_df, left_on='Name', right_on='proWandName', how='left')

    wand_grouped = wand_long.groupby(['emailAddress', 'Date', 'projectName']).agg({
        'Hours': 'sum'}).reset_index()

    # === Merge SAP and WAND

    # Force 'Date' column in both dataframes to be datetime64[ns]
    sap_grouped['Date'] = pd.to_datetime(sap_grouped['Date'])
    wand_grouped['Date'] = pd.to_datetime(wand_grouped['Date'])

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

    # === Save to Excel with formatting ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged.to_excel(writer, index=False, sheet_name='Comparison Report')

        workbook = writer.book
        worksheet = writer.sheets['Comparison Report']

        # === Format styles ===
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'center',
            'fg_color': '#DDEBF7', 'border': 1
        })
        number_format = workbook.add_format(
            {'num_format': '0.00', 'border': 1})
        text_format = workbook.add_format({'border': 1})
        highlight_format = workbook.add_format({
            'bg_color': '#FFD6D6', 'border': 1
        })
        date_format = workbook.add_format(
            {'num_format': 'dd-mmm-yyyy', 'border': 1})

        # Set column widths
        worksheet.set_column('A:A', 12)   # Date
        worksheet.set_column('B:B', 30)   # Email
        worksheet.set_column('C:C', 25)   # Full Name
        worksheet.set_column('D:D', 20)   # proWandName
        worksheet.set_column('E:G', 12)   # Hours & Delta

        # Apply header formatting
        for col_num, value in enumerate(merged.columns):
            worksheet.write(0, col_num, value, header_format)

        # Apply data formatting
        for row_num in range(1, len(merged) + 1):
            for col_num, col_name in enumerate(merged.columns):
                value = merged.iloc[row_num - 1, col_num]
                # Choose format
                if col_name == 'Date':
                    fmt = date_format
                elif col_name.startswith('Hours') or col_name == 'Delta':
                    fmt = number_format
                else:
                    fmt = text_format
                fmt = number_format if col_name.startswith(
                    'Hours') or col_name == 'Delta' else text_format

                # Highlight if delta ≠ 0
                if col_name == 'Delta' and value != 0:
                    fmt = highlight_format

                worksheet.write(row_num, col_num, value, fmt)

        # === Sheet 2: Consolidated Summary ===
        sap_summary = sap_long.groupby(['Email', 'Full Name', 'Date'])[
            'Hours'].sum().reset_index()
        wand_summary = wand_long.groupby(['emailAddress', 'Date'])[
            'Hours'].sum().reset_index()

        all_emails = set(sap_summary['Email']).union(
            set(wand_summary['emailAddress']))
        summary_data = []

        for email in all_emails:
            sap_dates = set(sap_summary[sap_summary['Email'] == email]['Date'])
            wand_dates = set(
                wand_summary[wand_summary['emailAddress'] == email]['Date'])

            extra_sap = sorted([d.strftime("%d")
                               for d in (sap_dates - wand_dates)])
            extra_wand = sorted([d.strftime("%d")
                                for d in (wand_dates - sap_dates)])

            sap_total = sap_summary[sap_summary['Email']
                                    == email]['Hours'].sum()
            wand_total = wand_summary[wand_summary['emailAddress'] == email]['Hours'].sum(
            )

            name = sap_summary[sap_summary['Email'] ==
                               email]['Full Name'].iloc[0] if email in sap_summary['Email'].values else ""
            team = mapping_df[mapping_df[email_col] ==
                              email]['projectName'].iloc[0] if email in mapping_df[email_col].values else "N/A"

            summary_data.append({
                "Email": email,
                "Full Name": name,
                "Days Only in SAP": extra_sap,
                "Days Only in WAND": extra_wand,
                "Total SAP Hours": round(sap_total, 2),
                "Total WAND Hours": round(wand_total, 2),
                "Hour Difference": round(sap_total - wand_total, 2),
                "Team": team
            })

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, index=False, sheet_name='Summary')

        # === Optional: Format the Summary sheet ===
        summary_ws = writer.sheets['Summary']
        summary_ws.set_column('A:A', 30)  # Email
        summary_ws.set_column('B:B', 25)  # Full Name
        summary_ws.set_column('C:D', 30)  # Extra Days
        summary_ws.set_column('E:G', 18)  # Totals & Diff
        summary_ws.set_column('H:H', 15)  # Team

        # Header format for summary
        for col_num, value in enumerate(summary_df.columns):
            summary_ws.write(0, col_num, value, header_format)

    output.seek(0)
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={"Content-Disposition": "attachment; filename=comparison_report.xlsx"}
    )
