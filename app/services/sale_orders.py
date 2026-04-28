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


def _get_sheet():
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds      = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client     = gspread.authorize(creds)
    sheet      = client.open_by_key(SPREADSHEET_ID)
    return sheet.worksheet(SHEET_TAB)


# ── DATE PARSING ──────────────────────────────────────────────────────────────

def _parse_order_date(raw: str) -> str:
    """
    Accepts:
      - dd/mm/yyyy hh:MM:ss  (Unicommerce bulk CSV format)
      - ISO 8601 strings
      - Falls back to current UTC time
    Returns ISO 8601 string expected by Unicommerce API.
    """
    raw = str(raw).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except ValueError:
            continue
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── GROUP ROWS INTO ORDERS ────────────────────────────────────────────────────

def _group_rows_into_orders(records: list[dict]) -> dict[str, dict]:
    """
    The CSV can have multiple rows per order (one per item).
    Groups them by Sales Order Code, building a single payload per order
    with multiple saleOrderItems.
    """
    orders: dict[str, dict] = {}

    for row in records:
        order_code = str(row.get("Sales Order Code*", "")).strip()
        if not order_code:
            continue

        # ── First time we see this order code: build the header ──
        if order_code not in orders:
            ship_addr_id  = str(row.get("Shipping Address Id", f"SHIP-{order_code}")).strip() or f"SHIP-{order_code}"
            bill_addr_id  = str(row.get("Billing Address Id",  ship_addr_id)).strip()          or ship_addr_id

            ship_name     = str(row.get("Shipping Address Name",  "")).strip()
            ship_line1    = str(row.get("Shipping Address Line 1","")).strip()
            ship_line2    = str(row.get("Shipping Address Line 2","")).strip()
            ship_city     = str(row.get("Shipping Address City",  "")).strip()
            ship_state    = str(row.get("Shipping Address State", "")).strip()
            ship_country  = str(row.get("Shipping Address Country","India")).strip() or "India"
            ship_pincode  = str(row.get("Shipping Address Pincode","")).strip()
            ship_phone    = str(row.get("Shipping Address Phone", "")).strip()

            bill_name     = str(row.get("Billing Address Name",   ship_name)).strip()   or ship_name
            bill_line1    = str(row.get("Billing Address Line 1", ship_line1)).strip()  or ship_line1
            bill_line2    = str(row.get("Billing Address Line 2", ship_line2)).strip()  or ship_line2
            bill_city     = str(row.get("Billing Address City",   ship_city)).strip()   or ship_city
            bill_state    = str(row.get("Billing Address State",  ship_state)).strip()  or ship_state
            bill_country  = str(row.get("Billing Address Country",ship_country)).strip()or ship_country
            bill_pincode  = str(row.get("Billing Address Pincode",ship_pincode)).strip()or ship_pincode
            bill_phone    = str(row.get("Billing Address Phone",  ship_phone)).strip()  or ship_phone

            mobile        = str(row.get("Notification Mobile", ship_phone)).strip()
            cod_raw       = str(row.get("COD*", "FALSE")).strip().upper()
            cod           = cod_raw in ("TRUE", "1", "YES")
            channel       = str(row.get("Channel", "CUSTOM")).strip() or "CUSTOM"
            display_code  = str(row.get("Display Sales Order Code", order_code)).strip() or order_code
            order_date    = _parse_order_date(row.get("Order Date as dd/mm/yyyy hh:MM:ss", ""))
            currency      = str(row.get("Currency Code", "INR")).strip() or "INR"
            facility_code = str(row.get("Facility Code", "")).strip()

            # Numeric order-level totals (optional in CSV, default 0)
            def _num(key): return float(row.get(key) or 0)

            orders[order_code] = {
                "_facility_code": facility_code,   # popped before sending to API
                "saleOrder": {
                    "code":                       order_code,
                    "displayOrderCode":           display_code,
                    "displayOrderDateTime":       order_date,
                    "customerName":               ship_name,
                    "notificationMobile":         mobile,
                    "channel":                    channel,
                    "cashOnDelivery":             cod,
                    "currencyCode":               currency,
                    "totalPrepaidAmount":         _num("Prepaid Amount"),
                    "totalCashOnDeliveryCharges": _num("COD Service Charges"),
                    "totalDiscount":              _num("Discount"),
                    "totalShippingCharges":       _num("Order Total Shipping Charges"),
                    "totalGiftWrapCharges":       _num("Gift Wrap Charges"),
                    "totalStoreCredit":           _num("Store Credit"),
                    "addresses": [
                        {
                            "id":           ship_addr_id,
                            "name":         ship_name,
                            "addressLine1": ship_line1,
                            "addressLine2": ship_line2,
                            "city":         ship_city,
                            "state":        ship_state,
                            "country":      ship_country,
                            "pincode":      ship_pincode,
                            "phone":        ship_phone,
                        },
                    ],
                    "shippingAddress": {"referenceId": ship_addr_id},
                    "billingAddress":  {"referenceId": bill_addr_id},
                    "saleOrderItems":  [],
                },
            }

            # Add billing address as a separate entry only if it differs from shipping
            if bill_addr_id != ship_addr_id:
                orders[order_code]["saleOrder"]["addresses"].append({
                    "id":           bill_addr_id,
                    "name":         bill_name,
                    "addressLine1": bill_line1,
                    "addressLine2": bill_line2,
                    "city":         bill_city,
                    "state":        bill_state,
                    "country":      bill_country,
                    "pincode":      bill_pincode,
                    "phone":        bill_phone,
                })

        # ── Append line item (every row contributes one item) ──
        item_code     = str(row.get("Sale Order Item Code*", "")).strip()
        sku           = str(row.get("Item SKU Code*", "")).strip()
        shipping_meth = str(row.get("Shipping Method*", "STD")).strip() or "STD"
        quantity      = int(float(row.get("Quantity") or 1))
        packet_num    = int(float(row.get("Packet Number") or 1))
        gift_wrap     = str(row.get("Gift Wrap", "FALSE")).strip().upper() in ("TRUE", "1", "YES")
        on_hold       = str(row.get("On Hold",   "FALSE")).strip().upper() in ("TRUE", "1", "YES")

        def _item_num(key): return float(row.get(key) or 0)

        item = {
            "code":               item_code,
            "itemSku":            sku,
            "shippingMethodCode": shipping_meth,
            "facilityCode":       str(row.get("Facility Code", "")).strip(),
            "packetNumber":       packet_num,
            "giftWrap":           gift_wrap,
            "onHold":             on_hold,
            "totalPrice":         _item_num("Selling Price"),
            "sellingPrice":       _item_num("Selling Price"),
            "prepaidAmount":      _item_num("Prepaid Amount"),
            "discount":           _item_num("Discount"),
            "shippingCharges":    _item_num("Shipping Charges"),
            "storeCredit":        _item_num("Store Credit"),
            "giftWrapCharges":    _item_num("Gift Wrap Charges"),
            "quantity":           quantity,
        }

        # Optional fields — only include if present in the row
        voucher_code = str(row.get("Voucher Code", "")).strip()
        if voucher_code:
            item["voucherCode"]  = voucher_code
            item["voucherValue"] = _item_num("Voucher Value")

        tracking = str(row.get("Tracking Number", "")).strip()
        if tracking:
            item["trackingNumber"] = tracking

        orders[order_code]["saleOrder"]["saleOrderItems"].append(item)

    return orders


# ── STATUS COLUMN HELPER ──────────────────────────────────────────────────────

def _find_or_create_status_col(ws, all_values: list) -> int:
    raw_headers = all_values[0] if all_values else []

    if "Order Status" in raw_headers:
        return raw_headers.index("Order Status") + 1

    sheet_props = ws.spreadsheet.fetch_sheet_metadata()
    max_cols    = 26
    for s in sheet_props.get("sheets", []):
        props = s.get("properties", {})
        if props.get("title") == ws.title:
            max_cols = props.get("gridProperties", {}).get("columnCount", 26)
            break

    next_col = len(raw_headers) + 1
    if next_col > max_cols:
        ws.spreadsheet.batch_update({
            "requests": [{
                "appendDimension": {
                    "sheetId":   ws.id,
                    "dimension": "COLUMNS",
                    "length":    1,
                }
            }]
        })

    ws.update_cell(1, next_col, "Order Status")
    return next_col


# ── MAIN ENTRY POINT ──────────────────────────────────────────────────────────

def process_sale_orders() -> dict:
    ws         = _get_sheet()
    all_values = ws.get_all_values()

    if not all_values or len(all_values) < 2:
        return {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    raw_headers = all_values[0]

    # De-duplicate headers
    seen    = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    # Build per-row dicts and track which sheet row each belongs to
    # row_index[order_code] = earliest sheet row (used for status write-back)
    records:   list[dict] = []
    row_index: dict[str, int] = {}

    for sheet_row, row_values in enumerate(all_values[1:], start=2):
        padded = row_values + [""] * (len(headers) - len(row_values))
        row    = dict(zip(headers, padded))
        records.append(row)

        order_code = str(row.get("Sales Order Code*", "")).strip()
        if order_code and order_code not in row_index:
            row_index[order_code] = sheet_row

    # Find / create status column
    status_col = _find_or_create_status_col(ws, all_values)

    # Read current statuses for all known rows to avoid re-processing
    status_cache: dict[int, str] = {}
    for order_code, sheet_row in row_index.items():
        val = ws.cell(sheet_row, status_col).value or ""
        status_cache[sheet_row] = str(val).strip()

    # Group into orders (handles multi-row / multi-item orders)
    orders = _group_rows_into_orders(records)

    url     = f"{TENANT_URL}/services/rest/v1/oms/saleOrder/create"
    results = {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    for order_code, order_data in orders.items():
        sheet_row = row_index.get(order_code)

        # Skip already processed
        current_status = status_cache.get(sheet_row, "")
        if current_status in ("SUCCESS", "FAILED"):
            results["skipped"] += 1
            continue

        facility_code = order_data.pop("_facility_code", "")
        sale_order    = order_data  # now just {"saleOrder": {...}}

        if not facility_code:
            msg = "SKIPPED - missing Facility Code"
            if sheet_row:
                ws.update_cell(sheet_row, status_col, msg)
            results["skipped"] += 1
            continue

        if not sale_order["saleOrder"]["saleOrderItems"]:
            msg = "SKIPPED - no items"
            if sheet_row:
                ws.update_cell(sheet_row, status_col, msg)
            results["skipped"] += 1
            continue

        try:
            data = api_post(url, sale_order, facility=facility_code)

            if data.get("successful"):
                uc_code = data.get("saleOrderDetailDTO", {}).get("code", order_code)
                status  = f"SUCCESS - {uc_code}"
                results["success"] += 1
                print(f"  ✔ Order {order_code} → {uc_code}")
            else:
                errors = data.get("errors", [])
                msg    = errors[0].get("description") if errors else data.get("message", "Unknown error")
                status = f"FAILED - {msg}"
                results["failed"] += 1
                results["errors"].append({"order_id": order_code, "error": msg})
                print(f"  ✗ Order {order_code} failed: {msg}")

        except Exception as e:
            status = f"ERROR - {str(e)[:120]}"
            results["failed"] += 1
            results["errors"].append({"order_id": order_code, "error": str(e)})
            print(f"  ✗ Order {order_code} exception: {e}")

        if sheet_row:
            ws.update_cell(sheet_row, status_col, status)

    return results