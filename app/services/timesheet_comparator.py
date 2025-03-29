import pandas as pd
from app.utils.load_mapping_file import load_mapping_file
from app.utils.load_sap_file import load_sap_file
from app.utils.process_sap_data import process_sap_data
from app.utils.load_wand_file import load_wand_file
from app.utils.process_wand_data import process_wand_data
from app.utils.excel_writer import write_report_to_excel


async def process_timesheets(sap_file, wand_file, mapping_file):
    # === Validate Inputs ===
    sap_path = sap_file.file if sap_file else None
    wand_path = wand_file.file if wand_file else None
    mapping_path = mapping_file.file if mapping_file else None

    if not (sap_path and wand_path and mapping_path):
        raise ValueError("❌ All files (SAP, WAND, Mapping) must be provided.")

    # === Load Files ===
    mapping_df, email_col = load_mapping_file(mapping_path)
    if mapping_df.empty or not email_col:
        raise ValueError(
            "❌ Failed to load mapping file or detect email column.")

    sap_raw_data, dates = load_sap_file(sap_path)
    if sap_raw_data.empty:
        raise ValueError("❌ Failed to load SAP file.")
    sap_grouped, sap_long = process_sap_data(
        sap_raw_data, dates, mapping_df, email_col)

    wand_df = load_wand_file(wand_path)
    if wand_df.empty:
        raise ValueError("❌ Failed to load WAND file.")
    wand_grouped = process_wand_data(wand_df, mapping_df)

    # === Merge Process ===
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

    # === Reorder Columns ===
    desired_order = [
        'Date', 'Email', 'Full Name', 'proWandName',
        'projectName', 'Hours_sap', 'Hours_wand', 'Delta'
    ]
    merged = merged[desired_order]

    # === Final Excel Report Generation ===
    output = write_report_to_excel(merged, mapping_df, email_col)
    return output
