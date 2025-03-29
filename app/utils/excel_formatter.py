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


def format_chart_sheet(workbook, sheets, summary_df):
    worksheet = workbook.add_worksheet('Charts')

    # Add chart 1: Bar chart of SAP vs WAND hours
    bar_chart = workbook.add_chart({'type': 'column'})

    row_count = len(summary_df)

    bar_chart.add_series({
        'name': 'SAP Hours',
        'categories': ['Summary', 1, 0, row_count, 0],  # Email column
        'values':     ['Summary', 1, 4, row_count, 4],  # Total SAP Hours
        'fill':       {'color': '#5B9BD5'},
    })
    bar_chart.add_series({
        'name': 'WAND Hours',
        'categories': ['Summary', 1, 0, row_count, 0],
        'values':     ['Summary', 1, 5, row_count, 5],  # Total WAND Hours
        'fill':       {'color': '#ED7D31'},
    })

    bar_chart.set_title({'name': 'SAP vs WAND Hours'})
    bar_chart.set_x_axis({'name': 'Email'})
    bar_chart.set_y_axis({'name': 'Hours'})
    bar_chart.set_style(11)

    worksheet.insert_chart('B2', bar_chart, {'x_scale': 1.5, 'y_scale': 1.5})

    # Add chart 2: Line chart of Hour Difference
    line_chart = workbook.add_chart({'type': 'line'})

    line_chart.add_series({
        'name': 'Hour Difference',
        'categories': ['Summary', 1, 0, row_count, 0],  # Email column
        'values':     ['Summary', 1, 6, row_count, 6],  # Hour Difference
        'line':       {'color': 'red'},
    })

    line_chart.set_title({'name': 'Hour Difference Per Employee'})
    line_chart.set_x_axis({'name': 'Email'})
    line_chart.set_y_axis({'name': 'Difference'})
    line_chart.set_style(10)

    worksheet.insert_chart('B20', line_chart, {'x_scale': 1.5, 'y_scale': 1.5})
