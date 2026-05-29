import pandas as pd

try:
    import gspread
    from google.oauth2 import service_account
    _GSPREAD_OK = True
except ImportError:
    _GSPREAD_OK = False

_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
_sp_cache: dict = {}


def is_configured() -> bool:
    if not _GSPREAD_OK:
        return False
    try:
        import streamlit as st
        creds = dict(st.secrets.get("gcp_service_account", {}))
        sheet_id = str(st.secrets.get("SPREADSHEET_ID", ""))
        return bool(sheet_id and creds.get("type") == "service_account")
    except Exception:
        return False


def _open():
    """Return a gspread Spreadsheet object, reusing a cached connection."""
    try:
        import streamlit as st
        creds_dict = dict(st.secrets.get("gcp_service_account", {}))
        sheet_id = str(st.secrets.get("SPREADSHEET_ID", ""))
        if not creds_dict or not sheet_id:
            return None
        cache_key = creds_dict.get("client_email", "") + "|" + sheet_id
        if cache_key in _sp_cache:
            return _sp_cache[cache_key]
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=_SCOPES)
        gc = gspread.authorize(creds)
        sp = gc.open_by_key(sheet_id)
        _sp_cache[cache_key] = sp
        return sp
    except Exception:
        return None


def read_sheet(name: str) -> pd.DataFrame:
    """Read a worksheet and return a DataFrame (all values as strings)."""
    try:
        sp = _open()
        if sp is None:
            return pd.DataFrame()
        ws = sp.worksheet(name)
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame()
        headers = values[0]
        rows = values[1:]
        if not rows:
            return pd.DataFrame(columns=headers)
        return pd.DataFrame(rows, columns=headers).fillna("")
    except Exception:
        return pd.DataFrame()


def write_all(sheets: dict) -> bool:
    """Write multiple sheets in one connection. sheets = {sheet_name: DataFrame}."""
    try:
        sp = _open()
        if sp is None:
            return False
        for name, df in sheets.items():
            ws = sp.worksheet(name)
            ws.clear()
            data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
            ws.update('A1', data)
        return True
    except Exception:
        return False
