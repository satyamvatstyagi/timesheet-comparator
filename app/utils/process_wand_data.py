import pandas as pd


def process_wand_data(wand_df, mapping_df):
    try:
        wand_long = pd.melt(
            wand_df,
            id_vars=['Name'],
            var_name='Date',
            value_name='Hours'
        )
        wand_long['Date'] = pd.to_datetime(wand_long['Date'], errors='coerce')
        wand_long['Hours'] = pd.to_numeric(wand_long['Hours'], errors='coerce')
        wand_long = wand_long.dropna(subset=['Date', 'Hours'])

        wand_long = wand_long[wand_long['Name'].isin(
            mapping_df['proWandName'])]
        wand_long = wand_long.merge(
            mapping_df, left_on='Name', right_on='proWandName', how='left')

        wand_grouped = wand_long.groupby(
            ['emailAddress', 'Date', 'projectName']
        ).agg({'Hours': 'sum'}).reset_index()

        return wand_grouped

    except Exception as e:
        raise RuntimeError(f"❌ Failed to process WAND data: {e}")
