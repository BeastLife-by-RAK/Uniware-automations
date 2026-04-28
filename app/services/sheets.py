import json
import os
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_ID   = os.getenv("GOOGLE_SHEETS_ID")
CREDENTIALS_JSON = os.getenv("GOOGLE_SHEETS_CREDENTIALS")  # full JSON string

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

FACILITY_SHORTNAME = {
    "Emiza_B2C_BLR":    "BLR",
    "Emiza_B2C_GGN":    "GGN",
    "Emiza_B2C_Mumbai": "MUM",
    "Emiza_B2C_WB":     "WB",
}


def _get_client() -> gspread.Client:
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds      = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def push_inventory_to_sheets(records: list[dict]) -> dict:
    """
    Groups records by facility and writes each into a dated tab.
    Tab name format: GGN-2026-04-27
    Returns summary of tabs created.
    """
    client        = _get_client()
    spreadsheet   = client.open_by_key(SPREADSHEET_ID)
    today         = datetime.now().strftime("%Y-%m-%d")
    summary       = {}

    # Group records by facility
    by_facility: dict[str, list[dict]] = {}
    for r in records:
        code = r.get("facilityCode", "UNKNOWN")
        by_facility.setdefault(code, []).append(r)

    for facility_code, rows in by_facility.items():
        short     = FACILITY_SHORTNAME.get(facility_code, facility_code)
        tab_name  = f"{short}-{today}"

        # Delete existing tab with same name if exists (re-run same day)
        try:
            existing = spreadsheet.worksheet(tab_name)
            spreadsheet.del_worksheet(existing)
            print(f"  Deleted existing tab: {tab_name}")
        except gspread.exceptions.WorksheetNotFound:
            pass

        # Create fresh tab
        ws = spreadsheet.add_worksheet(title=tab_name, rows=len(rows) + 5, cols=15)

        # Headers
        headers = [
            "SKU Code",
            "Facility Code",
            "Location",
            "Sellable Qty",
            "Open Sale",
            "Open Purchase",
            "Putaway Pending",
            "Blocked Qty",
            "Pending Stock Transfer",
            "Vendor Inventory",
            "Virtual Inventory",
            "Pending Assessment",
            "Bad Inventory",
            "Inventory Not Synced",
            "Batch Recall Qty",
        ]

        # Build all data rows
        data = [headers]
        for r in rows:
            data.append([
                r.get("itemTypeSKU"),
                r.get("facilityCode"),
                r.get("facilityLocation"),
                r.get("inventory", 0),
                r.get("openSale", 0),
                r.get("openPurchase", 0),
                r.get("putawayPending", 0),
                r.get("inventoryBlocked", 0),
                r.get("pendingStockTransfer", 0),
                r.get("vendorInventory", 0),
                r.get("virtualInventory", 0),
                r.get("pendingInventoryAssessment", 0),
                r.get("badInventory", 0),
                r.get("inventoryNotSynced", 0),
                r.get("batchRecallQuantity", 0),
            ])

        # Write all at once (single API call per facility)
        ws.update("A1", data)

        # Format header row — bold + background
        ws.format("A1:O1", {
            "backgroundColor": {"red": 0.122, "green": 0.306, "blue": 0.475},
            "textFormat":      {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
        })

        # Freeze header row
        spreadsheet.batch_update({
            "requests": [{
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": ws.id,
                        "gridProperties": {"frozenRowCount": 1}
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            }]
        })

        print(f"  ✓ Written {len(rows)} rows to tab: {tab_name}")
        summary[tab_name] = len(rows)

    return summary