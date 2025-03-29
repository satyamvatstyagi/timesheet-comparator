import pandas as pd
from io import BytesIO
from app.utils.excel_formatter import (
    format_comparison_sheet,
    format_summary_sheet,
    write_summary_sheet
)


async def process_timesheets(sap_file, wand_file, mapping_file):
    # === Load mapping file ===
    mapping_df = pd.read_csv(mapping_file.file)
    mapping_df.columns = mapping_df.columns.str.strip()
    email_col = next(
        (col for col in mapping_df.columns if 'email' in col.lower()), None)
    if email_col is None:
        raise ValueError("'email' column not found in mapping.csv")

    # === Load SAP ===
    sap_raw = pd.read_excel(sap_file.file, header=None, skiprows=19)
    date_row = sap_raw.iloc[0, 6:].tolist()
    dates = pd.to_datetime(date_row, errors='coerce')
    sap_data = sap_raw.iloc[2:].reset_index(drop=True)

    emails = sap_data.iloc[:, 4]
    full_names = sap_data.iloc[:, 3]
    hours_data = sap_data.iloc[:, 6:].applymap(
        lambda x: float(str(x).replace(" H", "").strip()
                        ) if pd.notnull(x) else 0
    )

    sap_long = pd.melt(hours_data.reset_index(drop=True),
                       var_name="DayIndex", value_name="Hours")
    sap_long['Email'] = emails.repeat(
        hours_data.shape[1]).reset_index(drop=True)
    sap_long['Full Name'] = full_names.repeat(
        hours_data.shape[1]).reset_index(drop=True)
    sap_long['Date'] = list(dates) * len(sap_data)
    sap_long = sap_long[['Email', 'Full Name', 'Date', 'Hours']].dropna(subset=[
                                                                        'Date', 'Hours'])

    sap_grouped = sap_long.groupby(['Email', 'Full Name', 'Date']).agg({
        'Hours': 'sum'}).reset_index()
    sap_grouped = sap_grouped.merge(
        mapping_df, left_on='Email', right_on=email_col, how='left')
    sap_grouped.rename(columns={email_col: 'emailAddress'}, inplace=True)
    sap_grouped.drop(columns=['Email'], inplace=True)

    # === Load WAND ===
    wand_df = pd.read_excel(wand_file.file, engine='xlrd')
    wand_long = pd.melt(
        wand_df, id_vars=['Name'], var_name='Date', value_name='Hours')
    wand_long['Date'] = pd.to_datetime(wand_long['Date'], errors='coerce')
    wand_long['Hours'] = pd.to_numeric(wand_long['Hours'], errors='coerce')
    wand_long = wand_long.dropna(subset=['Date', 'Hours'])

    wand_long = wand_long[wand_long['Name'].isin(mapping_df['proWandName'])]
    wand_long = wand_long.merge(
        mapping_df, left_on='Name', right_on='proWandName', how='left')

    wand_grouped = wand_long.groupby(['emailAddress', 'Date', 'projectName']).agg({
        'Hours': 'sum'
    }).reset_index()

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
    merged.rename(columns={'emailAddress': 'Email'}, inplace=True)

    desired_order = ['Date', 'Email', 'Full Name', 'proWandName',
                     'projectName', 'Hours_sap', 'Hours_wand', 'Delta']
    merged = merged[desired_order]

    # === Write to Excel ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd-mmm-yyyy') as writer:
        merged.to_excel(writer, index=False, sheet_name='Comparison Report')
        worksheet = writer.sheets['Comparison Report']
        format_comparison_sheet(writer.book, worksheet, merged)

        summary_df = write_summary_sheet(writer, merged, mapping_df, email_col)
        summary_ws = writer.sheets['Summary']
        format_summary_sheet(writer.book, summary_ws, summary_df)

    output.seek(0)
    return output
