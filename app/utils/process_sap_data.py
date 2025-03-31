import pandas as pd


def process_sap_data(sap_data, dates, mapping_df, email_col):
    try:
        print("📥 Starting SAP data processing...")

        emails = sap_data.iloc[:, 4]
        full_names = sap_data.iloc[:, 3]
        print(f"📧 Sample Emails: {emails.head(3).tolist()}")
        print(f"👤 Sample Names: {full_names.head(3).tolist()}")

        # Extract only hour columns (assuming from col 5 onwards)
        hours_data = sap_data.iloc[:, 5:]
        print(f"🕒 Raw Hours Data (head):\n{hours_data.head(3)}")

        # Clean and convert hour values
        hours_data = hours_data.applymap(
            lambda x: float(str(x).replace(" H", "").strip()
                            ) if pd.notnull(x) else 0
        )
        print(f"✅ Cleaned Hours Data (head):\n{hours_data.head(3)}")

        # Validate dates match columns
        if len(dates) != hours_data.shape[1]:
            raise ValueError(
                "❌ Dates extracted do not match SAP hours columns")
        print(f"📅 Dates extracted: {dates}")

        # Assign column names to be dates
        hours_data.columns = pd.to_datetime(dates)

        # Melt (stack) into long format
        sap_long = hours_data.stack().reset_index()
        sap_long.columns = ['RowIndex', 'Date', 'Hours']
        print(f"🔁 Long Format SAP Hours (head):\n{sap_long.head(5)}")

        # Add corresponding email and name by using RowIndex
        sap_long['Email'] = sap_data.iloc[sap_long['RowIndex'], 4].values
        sap_long['Full Name'] = sap_data.iloc[sap_long['RowIndex'], 3].values

        # Clean types
        sap_long['Date'] = pd.to_datetime(sap_long['Date'], errors='coerce')
        sap_long['Hours'] = pd.to_numeric(
            sap_long['Hours'], errors='coerce').fillna(0)

        # Drop rows where email is missing (empty rows)
        sap_long = sap_long.dropna(subset=['Email'])

        print(f"📊 Final SAP Long Format Data (head):\n{sap_long.head(5)}")

        # Group by employee and date
        sap_grouped = sap_long.groupby(['Email', 'Full Name', 'Date']).agg({
            'Hours': 'sum'
        }).reset_index()
        print(f"📦 Grouped SAP Data:\n{sap_grouped.head(5)}")

        # Merge mapping info
        sap_grouped = sap_grouped.merge(
            mapping_df, left_on='Email', right_on=email_col, how='left')
        sap_grouped.rename(columns={email_col: 'emailAddress'}, inplace=True)
        sap_grouped.drop(columns=['Email'], inplace=True)

        print("✅ SAP data processing completed successfully.\n")
        return sap_grouped, sap_long

    except Exception as e:
        raise RuntimeError(f"❌ Failed to process SAP data: {e}")
