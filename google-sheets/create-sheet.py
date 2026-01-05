from __future__ import annotations

from typing import Dict, List, Tuple
from google.oauth2 import service_account
from googleapiclient.discovery import build


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "service_account.json"

SPREADSHEET_TITLE = "צומחים_ביחד_דאטה"
# Set this to an existing spreadsheet ID, or leave as None to create a new one
EXISTING_SPREADSHEET_ID = "1EDpuvzbFF1pOEuMU9wC5mpUoNrmet-pu3GB6FpECt9I"

SHEETS: Dict[str, List[str]] = {
    "יומן_שיחות": [
        "מזהה_שיחה",
        "תאריך_ושעה",
        "טלפון",
        "תפקיד_במיזם",
        "מטרת_פנייה",
        "סטטוס_שיחה",
        "הסכמה_לשמירת_פרטים",
        "הערות",
        "קישור_להקלטה",
        "תמלול",
    ],
    "מצמיחים": [
        "מזהה_מצמיח",
        "שם_מלא",
        "טלפון",
        "כתובת_ברעננה",
        "ימים_מועדפים",
        "חלון_זמן",
        "הערות_ניידות",
        "מקור_הרישום",
        "שם_הארגון_התומך",
        "שם_נציג_הארגון",
        "סטטוס",
        "אחראי_אדמיניסטרטיבי",
        "תאריך_יצירה",
        "תאריך_עדכון_אחרון",
        "מזהה_מתנדב_לביקורי_בית",
    ],
    "מתנדבים_לביקורי_בית": [
        "מזהה_מתנדב",
        "שם_מלא",
        "טלפון",
        "זמינות_ברעננה",
        "תדירות_זמינות",
        "ימים_מועדפים",
        "חלון_זמן",
        "יש_רכב",
        "סטטוס",
        "אחראי_אדמיניסטרטיבי",
        "מזהה_מצמיח",
        "תאריך_יצירה",
    ],
    "נציגי_ארגון_תומך": [
        "מזהה_נציג_ארגון_תומך",
        "שם_מלא",
        "ארגון",
        "תפקיד",
        "טלפון",
        "אימייל",
        "סטטוס",
        "תאריך_יצירה",
    ],
    "שיוכים": [
        "מזהה_שיוך",
        "מזהה_מצמיח",
        "מזהה_מתנדב",
        "סטטוס_שיוך",
        "תאריך_שיוך",
        "תדירות_מפגשים",
        "הערות_אדמיניסטרטיביות",
    ],
    "רשימות_בחירה": [
        "תפקיד_במיזם",
        "חלון_זמן",
        "סטטוס_כללי",
        "סטטוס_שיחה",
        "סטטוס_שיוך",
        "תדירות_מפגשים",
        "מקור_הרישום",
        "יש_רכב",
        "מטרת_פנייה",
    ],
}

PICKLISTS: Dict[str, List[str]] = {
    "תפקיד_במיזם": ["מצמיח", "מתנדב לביקורי בית", "נציג ארגון תומך", "לא זוהה"],
    "חלון_זמן": ["בוקר", "צהריים", "ערב"],
    "סטטוס_כללי": ["חדש", "בבדיקה", "שובץ", "פעיל", "מושהה", "הסתיים"],
    "סטטוס_שיחה": ["הושלמה", "נותקה", "נכשלה"],
    "סטטוס_שיוך": ["מוצע", "מאושר", "פעיל", "הסתיים"],
    "תדירות_מפגשים": ["שבועי"],
    "מקור_הרישום": ["פנייה_עצמית", "נציג_ארגון_תומך"],
    "יש_רכב": ["כן", "לא"],
    "מטרת_פנייה": ["רישום", "הפניה", "עדכון", "אחר"],
}


def col_letter(n: int) -> str:
    """1-indexed column number -> A, B, ..., AA"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def create_spreadsheet(service) -> Tuple[str, Dict[str, int]]:
    if EXISTING_SPREADSHEET_ID:
        # Use existing spreadsheet
        spreadsheet_id = EXISTING_SPREADSHEET_ID
        # Fetch existing sheets
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing_sheet_names = {sh["properties"]["title"] for sh in meta["sheets"]}
        sheet_ids: Dict[str, int] = {}
        
        # Map existing sheets
        for sh in meta["sheets"]:
            sheet_ids[sh["properties"]["title"]] = sh["properties"]["sheetId"]
        
        # Add missing sheets
        requests = []
        for name in SHEETS.keys():
            if name not in existing_sheet_names:
                requests.append({"addSheet": {"properties": {"title": name}}})
        
        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": requests}
            ).execute()
            # Re-fetch to get new sheet IDs
            meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            for sh in meta["sheets"]:
                sheet_ids[sh["properties"]["title"]] = sh["properties"]["sheetId"]
        
        return spreadsheet_id, sheet_ids
    
    # Create new spreadsheet
    body = {"properties": {"title": SPREADSHEET_TITLE}}
    resp = service.spreadsheets().create(body=body).execute()
    spreadsheet_id = resp["spreadsheetId"]

    # Fetch initial sheetId
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    first_sheet = meta["sheets"][0]
    first_sheet_id = first_sheet["properties"]["sheetId"]
    sheet_ids: Dict[str, int] = {}

    # Rename default sheet to the first desired sheet name
    first_name = list(SHEETS.keys())[0]
    requests = [{
        "updateSheetProperties": {
            "properties": {"sheetId": first_sheet_id, "title": first_name},
            "fields": "title",
        }
    }]

    # Add the rest
    for name in list(SHEETS.keys())[1:]:
        requests.append({"addSheet": {"properties": {"title": name}}})

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()

    # Re-fetch to map names -> sheetId
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sh in meta["sheets"]:
        sheet_ids[sh["properties"]["title"]] = sh["properties"]["sheetId"]

    return spreadsheet_id, sheet_ids


def set_headers_and_picklists(service, spreadsheet_id: str):
    # Set headers for each sheet (row 1)
    data = []
    for sheet_name, headers in SHEETS.items():
        rng = f"'{sheet_name}'!A1:{col_letter(len(headers))}1"
        data.append({"range": rng, "values": [headers]})

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()

    # Fill picklists in רשימות_בחירה columns under header (start row 2)
    pick_headers = SHEETS["רשימות_בחירה"]
    data = []
    for idx, header in enumerate(pick_headers, start=1):
        values = PICKLISTS.get(header, [])
        if not values:
            continue
        col = col_letter(idx)
        rng = f"'רשימות_בחירה'!{col}2:{col}{len(values)+1}"
        data.append({"range": rng, "values": [[v] for v in values]})

    if data:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": "RAW", "data": data},
        ).execute()


def add_data_validations(service, spreadsheet_id: str, sheet_ids: Dict[str, int]):
    """
    Adds dropdown validations using values from רשימות_בחירה sheet.
    Since cross-sheet range references with Hebrew names can be problematic,
    we read the values first and use ONE_OF_LIST instead.
    """
    # Ensure רשימות_בחירה sheet exists
    if "רשימות_בחירה" not in sheet_ids:
        raise ValueError("רשימות_בחירה sheet must exist before adding data validations")
    
    # Read picklist values from the sheet
    picklist_headers = SHEETS["רשימות_בחירה"]
    picklist_data = {}
    
    for idx, header in enumerate(picklist_headers, start=1):
        col = col_letter(idx)
        rng = f"'רשימות_בחירה'!{col}2:{col}200"
        try:
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=rng
            ).execute()
            values = result.get('values', [])
            # Flatten list of lists to list of strings
            picklist_data[header] = [v[0] if v and len(v) > 0 else '' for v in values if v and len(v) > 0 and v[0]]
        except Exception as e:
            # Fallback to PICKLISTS dict if reading fails
            picklist_data[header] = PICKLISTS.get(header, [])
    
    # Map picklist header -> column index (0-based)
    pick_col_map = {
        "תפקיד_במיזם": 0,
        "חלון_זמן": 1,
        "סטטוס_כללי": 2,
        "סטטוס": 2,            # alias for סטטוס_כללי
        "סטטוס_שיחה": 3,
        "סטטוס_שיוך": 4,
        "תדירות_מפגשים": 5,
        "מקור_הרישום": 6,
        "יש_רכב": 7,
        "מטרת_פנייה": 8,
    }

    # Helper to create a dropdown rule using a list of values
    def dv_rule(header_name: str) -> dict:
        values = picklist_data.get(header_name, PICKLISTS.get(header_name, []))
        if not values:
            return None
        return {
            "condition": {
                "type": "ONE_OF_LIST",
                "values": [{"userEnteredValue": v} for v in values if v],
            },
            "showCustomUi": True,
            "strict": True,
        }

    # Target ranges: apply to rows 2..1000 in relevant columns
    requests = []

    # Helper to apply validation if rule exists
    def apply(sheet_name: str, col_index_zero: int, header_name: str):
        rule = dv_rule(header_name)
        if rule is None:
            return
        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_ids[sheet_name],
                    "startRowIndex": 1,      # row 2
                    "endRowIndex": 1000,     # up to row 1000
                    "startColumnIndex": col_index_zero,
                    "endColumnIndex": col_index_zero + 1,
                },
                "rule": rule,
            }
        })

    # יומן_שיחות
    # Columns: תפקיד_במיזם (D=3), מטרת_פנייה (E=4), סטטוס_שיחה (F=5)
    apply("יומן_שיחות", 3, "תפקיד_במיזם")
    apply("יומן_שיחות", 4, "מטרת_פנייה")
    apply("יומן_שיחות", 5, "סטטוס_שיחה")

    # מצמיחים: חלון_זמן (F=5), מקור_הרישום (H=7), סטטוס_כללי (K=10)
    apply("מצמיחים", 5, "חלון_זמן")
    apply("מצמיחים", 7, "מקור_הרישום")
    apply("מצמיחים", 10, "סטטוס_כללי")

    # מתנדבים_לביקורי_בית: תדירות_זמינות (E=4), חלון_זמן (G=6), יש_רכב (H=7), סטטוס_כללי (I=8)
    apply("מתנדבים_לביקורי_בית", 4, "תדירות_מפגשים")
    apply("מתנדבים_לביקורי_בית", 6, "חלון_זמן")
    apply("מתנדבים_לביקורי_בית", 7, "יש_רכב")
    apply("מתנדבים_לביקורי_בית", 8, "סטטוס_כללי")

    # נציגי_ארגון_תומך: סטטוס_כללי (G=6)
    apply("נציגי_ארגון_תומך", 6, "סטטוס_כללי")

    # שיוכים: סטטוס_שיוך (D=3), תדירות_מפגשים (F=5)
    apply("שיוכים", 3, "סטטוס_שיוך")
    apply("שיוכים", 5, "תדירות_מפגשים")

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()


def freeze_and_format(service, spreadsheet_id: str, sheet_ids: Dict[str, int]):
    # Freeze header row, set RTL, and bold header row (basic)
    requests = []
    for name, sid in sheet_ids.items():
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sid,
                    "gridProperties": {"frozenRowCount": 1},
                    "rightToLeft": True,
                },
                "fields": "gridProperties.frozenRowCount,rightToLeft",
            }
        })
        # Bold header row
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sid,
                    "startRowIndex": 0,
                    "endRowIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {"textFormat": {"bold": True}}
                },
                "fields": "userEnteredFormat.textFormat.bold"
            }
        })

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()


def main():
    service = build_service()

    if EXISTING_SPREADSHEET_ID:
        print(f"📝 Using existing spreadsheet: {EXISTING_SPREADSHEET_ID}")
    else:
        print("🆕 Creating new spreadsheet...")

    spreadsheet_id, sheet_ids = create_spreadsheet(service)
    set_headers_and_picklists(service, spreadsheet_id)
    add_data_validations(service, spreadsheet_id, sheet_ids)
    freeze_and_format(service, spreadsheet_id, sheet_ids)

    print("✅ Spreadsheet configured successfully!")
    print("Spreadsheet ID:", spreadsheet_id)
    print("Open it here:")
    print(f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}")


if __name__ == "__main__":
    main()