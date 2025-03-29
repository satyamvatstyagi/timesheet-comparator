import pandas as pd


def process_sap_data(sap_data, dates, mapping_df, email_col):
    try:
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

        return sap_grouped, sap_long

    except Exception as e:
        raise RuntimeError(f"❌ Failed to process SAP data: {e}")
