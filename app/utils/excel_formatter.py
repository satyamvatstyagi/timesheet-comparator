# app/utils/excel_formatter.py

import pandas as pd


def format_comparison_sheet(workbook, worksheet, merged_df):
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'center',
        'fg_color': '#DDEBF7', 'border': 1
    })
    number_format = workbook.add_format({'num_format': '0.00', 'border': 1})
    text_format = workbook.add_format({'border': 1})
    highlight_format = workbook.add_format(
        {'bg_color': '#FFD6D6', 'border': 1})
    date_format = workbook.add_format(
        {'num_format': 'dd-mmm-yyyy', 'border': 1})

    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:G', 12)

    for col_num, value in enumerate(merged_df.columns):
        worksheet.write(0, col_num, value, header_format)

    for row_num in range(1, len(merged_df) + 1):
        for col_num, col_name in enumerate(merged_df.columns):
            value = merged_df.iloc[row_num - 1, col_num]
            if col_name == 'Date':
                fmt = date_format
            elif col_name.startswith('Hours') or col_name == 'Delta':
                fmt = highlight_format if col_name == 'Delta' and value != 0 else number_format
            else:
                fmt = text_format
            worksheet.write(row_num, col_num, value, fmt)


def write_summary_sheet(writer, merged_df, mapping_df, email_col):
    summary_data = []

    for email in merged_df['Email'].dropna().unique():
        df = merged_df[merged_df['Email'] == email]

        # Non-zero days only
        sap_dates = set(df[df['Hours_sap'] > 0]['Date'])
        wand_dates = set(df[df['Hours_wand'] > 0]['Date'])

        extra_sap = sorted([d.strftime("%d")
                           for d in (sap_dates - wand_dates)])
        extra_wand = sorted([d.strftime("%d")
                            for d in (wand_dates - sap_dates)])

        total_sap = df['Hours_sap'].sum()
        total_wand = df['Hours_wand'].sum()
        delta = total_sap - total_wand

        full_name = df['Full Name'].iloc[0] if 'Full Name' in df else ""
        team = mapping_df[mapping_df[email_col] ==
                          email]['projectName'].iloc[0] if email in mapping_df[email_col].values else "N/A"

        summary_data.append({
            "Email": email,
            "Full Name": full_name,
            "Days Only in SAP": extra_sap,
            "Days Only in WAND": extra_wand,
            "Total SAP Hours": round(total_sap, 2),
            "Total WAND Hours": round(total_wand, 2),
            "Hour Difference": round(delta, 2),
            "Team": team
        })

    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, index=False, sheet_name='Summary')
    return summary_df


def format_summary_sheet(workbook, worksheet, summary_df):
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'center',
        'fg_color': '#DDEBF7', 'border': 1
    })

    worksheet.set_column('A:A', 30)  # Email
    worksheet.set_column('B:B', 25)  # Full Name
    worksheet.set_column('C:D', 30)  # Extra Days
    worksheet.set_column('E:G', 18)  # Totals & Diff
    worksheet.set_column('H:H', 15)  # Team

    for col_num, value in enumerate(summary_df.columns):
        worksheet.write(0, col_num, value, header_format)
