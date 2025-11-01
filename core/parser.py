import pandas as pd
import os
import logging
from datetime import datetime, timedelta

def find_column(df, keywords):
    for col in df.columns:
        col_clean = col.strip().lower().replace('\n', '').replace('\r', '')
        if any(keyword.lower() in col_clean for keyword in keywords):
            logging.info(f"Found column '{col}' matching keywords {keywords}")
            return col
    logging.warning(f"No column found matching keywords {keywords}")
    return None

def safe_parse_date(val):
    try:
        if pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            base_date = datetime(1899, 12, 30)
            return (base_date + timedelta(days=int(val))).date()
        if isinstance(val, pd.Timestamp):
            return val.date()
        return pd.to_datetime(val).date()
    except Exception as e:
        logging.error(f"Failed to parse date {val}: {e}")
        return None

def parse_timestamp(ts):
    try:
        parsed = pd.to_datetime(ts, errors='coerce')
        if pd.isna(parsed):
            logging.error(f"Invalid timestamp: {ts}")
            return pd.NaT
        return parsed.tz_localize(None)
    except Exception as e:
        logging.error(f"Failed to parse timestamp {ts}: {e}")
        return pd.NaT
