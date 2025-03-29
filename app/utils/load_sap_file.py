import pandas as pd


def load_sap_file(sap_file):
    try:
        sap_raw = pd.read_excel(sap_file, header=None, skiprows=19)
        date_row = sap_raw.iloc[0, 6:].tolist()
        dates = pd.to_datetime(date_row, errors='coerce')
        sap_data = sap_raw.iloc[2:].reset_index(drop=True)

        if sap_data.empty:
            raise ValueError("SAP data is empty after skipping rows.")

        return sap_data, dates
    except Exception as e:
        raise RuntimeError(f"❌ Failed to load SAP file: {e}")
