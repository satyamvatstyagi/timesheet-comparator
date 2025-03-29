import pandas as pd


def load_wand_file(wand_file):
    try:
        wand_df = pd.read_excel(wand_file, engine='xlrd')
        if wand_df.empty:
            raise ValueError("WAND file appears to be empty.")
        return wand_df
    except Exception as e:
        raise RuntimeError(f"❌ Failed to load WAND file: {e}")
