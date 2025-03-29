import pandas as pd


def load_mapping_file(mapping_file):
    try:
        df = pd.read_csv(mapping_file)
        df.columns = df.columns.str.strip()

        # Detect email column
        email_col = next(
            (col for col in df.columns if 'email' in col.lower()), None)
        if not email_col:
            raise ValueError("❌ 'email' column not found in mapping file")

        return df, email_col

    except Exception as e:
        raise RuntimeError(f"❌ Failed to load mapping file: {e}")
