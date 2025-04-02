import re
import pandas as pd
from datetime import datetime


def process_wand_data(wand_df, mapping_df):
    try:
        wand_long = pd.melt(
            wand_df,
            id_vars=['Name'],
            var_name='Date',
            value_name='Hours'
        )
        print(f"wand_long data: {wand_long.head(30)}")

        # Current year to append
        current_year = datetime.now().year

        # Detect DD-MMM pattern (e.g., "03-Mar", "15-Apr", etc.)
        def is_day_month_format(date_str):
            return bool(re.fullmatch(r"\d{1,2}-[A-Za-z]{3}", date_str))

        # Date column in wand_long
        wand_long['Date'] = wand_long['Date'].astype(str).str.strip()

        # Append current year where format is DD-MMM
        wand_long['Date'] = wand_long['Date'].apply(
            lambda x: f"{x}-{current_year}" if is_day_month_format(x) else x
        )

        # Now convert to datetime
        wand_long['Date'] = pd.to_datetime(wand_long['Date'], errors='coerce')
        print(f"wand_long data after date conversion: {wand_long.head(30)}")

        wand_long['Hours'] = pd.to_numeric(wand_long['Hours'], errors='coerce')
        print(f"wand_long data after hours conversion: {wand_long.head(30)}")
        wand_long = wand_long.dropna(subset=['Date', 'Hours'])
        print(f"wand_long data after dropping NA: {wand_long.head(30)}")

        wand_long = wand_long[wand_long['Name'].isin(
            mapping_df['proWandName'])]
        print(
            f"wand_long data after filtering by mapping_df: {wand_long.head(30)}")
        wand_long = wand_long.merge(
            mapping_df, left_on='Name', right_on='proWandName', how='left')
        print(
            f"wand_long data after merging with mapping_df: {wand_long.head(30)}")

        wand_grouped = wand_long.groupby(
            ['emailAddress', 'Date', 'projectName']
        ).agg({'Hours': 'sum'}).reset_index()
        print(f"wand_grouped data: {wand_grouped.head(30)}")

        return wand_grouped

    except Exception as e:
        raise RuntimeError(f"❌ Failed to process WAND data: {e}")
