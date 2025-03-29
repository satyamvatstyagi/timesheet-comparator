from io import BytesIO
import pandas as pd
from app.utils.excel_formatter import (
    format_comparison_sheet,
    format_summary_sheet,
    write_summary_sheet,
    format_chart_sheet
)


def write_report_to_excel(merged_df, mapping_df, email_col):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd-mmm-yyyy') as writer:
        # Sheet 1: Comparison Report
        merged_df.to_excel(writer, index=False, sheet_name='Comparison Report')
        worksheet = writer.sheets['Comparison Report']
        format_comparison_sheet(writer.book, worksheet, merged_df)

        # Sheet 2: Summary
        summary_df = write_summary_sheet(
            writer, merged_df, mapping_df, email_col)
        summary_ws = writer.sheets['Summary']
        format_summary_sheet(writer.book, summary_ws, summary_df)

        # Sheet 3: Charts
        format_chart_sheet(writer.book, writer.sheets, summary_df)

    output.seek(0)
    return output
