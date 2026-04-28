import os
import json
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials

from app.services.auth import api_post, TENANT_URL

# ── SHEET CONFIG ──────────────────────────────────────────────────────────────
SPREADSHEET_ID   = os.getenv("INFLUENCER_SHEET_ID")
SHEET_TAB        = os.getenv("INFLUENCER_SHEET_TAB", "Sheet1")
CREDENTIALS_JSON = os.getenv("GOOGLE_SHEETS_CREDENTIALS")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── FACILITY MAPPING ──────────────────────────────────────────────────────────
WAREHOUSE_TO_FACILITY = {
    "Gurgaon":   "Emiza_B2C_GGN",
    "Gurgoan":   "Emiza_B2C_GGN",   # handle typo in sheet
    "Bangalore": "Emiza_B2C_BLR",
    "Mumbai":    "Emiza_B2C_Mumbai",
    "Kolkata":   "Emiza_B2C_WB",
}


def _get_sheet():
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds      = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client     = gspread.authorize(creds)
    sheet      = client.open_by_key(SPREADSHEET_ID)
    return sheet.worksheet(SHEET_TAB)


def _build_order_payload(row: dict, row_number: int) -> dict:
    """Build Unicommerce sale order payload from a sheet row."""

    order_id     = str(row.get("Order ID", "")).strip()
    facility_raw = str(row.get("Near Warehouse", "")).strip()
    facility     = WAREHOUSE_TO_FACILITY.get(facility_raw, "Emiza_B2C_GGN")
    mobile       = str(row.get("*CustomerMobile", "")).strip()
    first_name   = str(row.get("*Customer First Name", "")).strip()
    last_name    = str(row.get("Customer Last Name", "")).strip()
    customer     = f"{first_name} {last_name}".strip()
    address1     = str(row.get("*Shipping Address Line 1", "")).strip()
    city         = str(row.get("*Shipping Address City", "")).strip()
    state        = str(row.get("*Shipping Address State", "")).strip()
    pincode      = str(row.get("*Shipping Address Postcode", "")).strip()
    sku          = str(row.get("*Master SKU", "")).strip()
    quantity     = int(row.get("*Product Quantity") or 1)
    order_date   = str(row.get("Order Place Date", datetime.now().isoformat())).strip()

    address_id   = "SHIP_ADDR_1"

    return {
        "saleOrder": {
            "code":                        order_id,
            "displayOrderCode":            order_id,
            "displayOrderDateTime":        order_date,
            "customerName":                customer,
            "notificationMobile":          mobile,
            "channel":                     "INFLUENCER",
            "cashOnDelivery":              False,
            "currencyCode":                "INR",
            "totalPrepaidAmount":          0,
            "totalCashOnDeliveryCharges":  0,
            "totalDiscount":               0,
            "totalShippingCharges":        0,
            "totalGiftWrapCharges":        0,
            "totalStoreCredit":            0,
            "addresses": [
                {
                    "id":           address_id,
                    "name":         customer,
                    "addressLine1": address1,
                    "city":         city,
                    "state":        state,
                    "country":      "India",
                    "pincode":      pincode,
                    "phone":        mobile,
                }
            ],
            "shippingAddress": {"referenceId": address_id},
            "billingAddress":  {"referenceId": address_id},
            "saleOrderItems": [
                {
                    "code":               f"{order_id}-1",
                    "itemSku":            sku,
                    "shippingMethodCode": "STD",
                    "facilityCode":       "",
                    "packetNumber":       1,
                    "giftWrap":           False,
                    "totalPrice":         0,
                    "sellingPrice":       0,
                    "prepaidAmount":      0,
                    "discount":           0,
                    "shippingCharges":    0,
                    "storeCredit":        0,
                    "giftWrapCharges":    0,
                }
            ],
        }
    }


def process_sale_orders() -> dict:
    ws      = _get_sheet()
    
    # Handle duplicate headers manually
    all_values = ws.get_all_values()
    if not all_values:
        return {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    raw_headers = all_values[0]
    
    # Rename duplicate headers by appending index
    seen    = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    # Build records manually
    records = []
    for row_values in all_values[1:]:
        # Pad row if shorter than headers
        padded = row_values + [""] * (len(headers) - len(row_values))
        records.append(dict(zip(headers, padded)))

    # Find or create Status column
    if "Order Status" not in raw_headers:
        ws.update_cell(1, len(raw_headers) + 1, "Order Status")
        status_col = len(raw_headers) + 1
    else:
        status_col = raw_headers.index("Order Status") + 1

    url     = f"{TENANT_URL}/services/rest/v1/oms/saleOrder/create"
    results = {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    for i, row in enumerate(records):
        row_number   = i + 2
        order_id     = str(row.get("Order ID", "")).strip()
        facility_raw = str(row.get("Near Warehouse", "")).strip()
        sku          = str(row.get("*Master SKU", "")).strip()
        status       = str(row.get("Order Status", "")).strip()

        if status in ("SUCCESS", "FAILED"):
            results["skipped"] += 1
            continue

        if not order_id or not sku or not facility_raw:
            ws.update_cell(row_number, status_col, "SKIPPED - missing data")
            results["skipped"] += 1
            continue

        facility = WAREHOUSE_TO_FACILITY.get(facility_raw)
        if not facility:
            ws.update_cell(row_number, status_col, f"SKIPPED - unknown warehouse: {facility_raw}")
            results["skipped"] += 1
            continue

        try:
            payload = _build_order_payload(row, row_number)
            data    = api_post(url, payload, facility=facility)

            if data.get("successful"):
                uc_code = data.get("saleOrderDetailDTO", {}).get("code", order_id)
                ws.update_cell(row_number, status_col, f"SUCCESS - {uc_code}")
                results["success"] += 1
                print(f"  ✓ Row {row_number} Order {order_id} → {uc_code}")
            else:
                errors = data.get("errors", [])
                msg    = errors[0].get("description") if errors else data.get("message", "Unknown error")
                ws.update_cell(row_number, status_col, f"FAILED - {msg}")
                results["failed"] += 1
                results["errors"].append({"order_id": order_id, "error": msg})
                print(f"  ✗ Row {row_number} Order {order_id} failed: {msg}")

        except Exception as e:
            ws.update_cell(row_number, status_col, f"ERROR - {str(e)[:100]}")
            results["failed"] += 1
            results["errors"].append({"order_id": order_id, "error": str(e)})
            print(f"  ✗ Row {row_number} Order {order_id} exception: {e}")

    return results