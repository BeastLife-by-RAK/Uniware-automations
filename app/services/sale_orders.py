import os
import json
from datetime import datetime

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


# ── GOOGLE SHEETS CLIENT ──────────────────────────────────────────────────────

def _get_sheet():
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds      = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client     = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_TAB)


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _s(row: dict, key: str, default: str = "") -> str:
    return row.get(key, default) or default


def _f(row: dict, key: str) -> float:
    try:
        return float(row.get(key) or 0)
    except (ValueError, TypeError):
        return 0.0


def _parse_order_date(raw: str) -> str:
    raw = raw.strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except ValueError:
            continue
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── GROUP ROWS INTO ORDERS ────────────────────────────────────────────────────

def _group_rows_into_orders(records: list[dict]) -> dict[str, dict]:
    """
    Groups sheet rows by Sales Order Code* into single API payloads.
    Rows with no Sales Order Code* are skipped.
    Multi-row orders (multiple SKUs) are merged into one payload.
    """
    orders: dict[str, dict] = {}

    for raw_row in records:
        # Strip every value once
        row = {k: str(v).strip() if v is not None else "" for k, v in raw_row.items()}

        order_code = row.get("Sales Order Code*", "")
        if not order_code:
            continue

        # ── Build order header on first encounter ─────────────────────────
        if order_code not in orders:
            ship_addr_id = _s(row, "Shipping Address Id") or f"SHIP-{order_code}"
            bill_addr_id = _s(row, "Billing Address Id")  or ship_addr_id

            ship_name    = _s(row, "Shipping Address Name")
            ship_line1   = _s(row, "Shipping Address Line 1")
            ship_line2   = _s(row, "Shipping Address Line 2")
            ship_city    = _s(row, "Shipping Address City")
            ship_state   = _s(row, "Shipping Address State")
            ship_country = _s(row, "Shipping Address Country") or "India"
            ship_pincode = _s(row, "Shipping Address Pincode")
            ship_phone   = _s(row, "Shipping Address Phone")

            bill_name    = _s(row, "Billing Address Name")    or ship_name
            bill_line1   = _s(row, "Billing Address Line 1")  or ship_line1
            bill_line2   = _s(row, "Billing Address Line 2")  or ship_line2
            bill_city    = _s(row, "Billing Address City")    or ship_city
            bill_state   = _s(row, "Billing Address State")   or ship_state
            bill_country = _s(row, "Billing Address Country") or ship_country
            bill_pincode = _s(row, "Billing Address Pincode") or ship_pincode
            bill_phone   = _s(row, "Billing Address Phone")   or ship_phone

            mobile       = _s(row, "Notification Mobile") or ship_phone
            cod          = _s(row, "COD*", "FALSE").upper() in ("TRUE", "1", "YES")
            channel      = _s(row, "Channel") or "CUSTOM"
            display_code = _s(row, "Display Sales Order Code") or order_code
            order_date   = _parse_order_date(_s(row, "Order Date as dd/mm/yyyy hh:MM:ss"))
            currency     = _s(row, "Currency Code") or "INR"
            facility     = _s(row, "Facility Code")

            addresses = [{
                "id":           ship_addr_id,
                "name":         ship_name,
                "addressLine1": ship_line1,
                "addressLine2": ship_line2,
                "city":         ship_city,
                "state":        ship_state,
                "country":      ship_country,
                "pincode":      ship_pincode,
                "phone":        ship_phone,
            }]

            if bill_addr_id != ship_addr_id:
                addresses.append({
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

            orders[order_code] = {
                "_facility_code": facility,
                "saleOrder": {
                    "code":                       order_code,
                    "displayOrderCode":           display_code,
                    "displayOrderDateTime":       order_date,
                    "customerName":               ship_name,
                    "notificationMobile":         mobile,
                    "channel":                    channel,
                    "cashOnDelivery":             cod,
                    "currencyCode":               currency,
                    "totalPrepaidAmount":         _f(row, "Prepaid Amount"),
                    "totalCashOnDeliveryCharges": _f(row, "COD Service Charges"),
                    "totalDiscount":              _f(row, "Discount"),
                    "totalShippingCharges":       _f(row, "Order Total Shipping Charges"),
                    "totalGiftWrapCharges":       _f(row, "Gift Wrap Charges"),
                    "totalStoreCredit":           _f(row, "Store Credit"),
                    "addresses":                  addresses,
                    "shippingAddress":            {"referenceId": ship_addr_id},
                    "billingAddress":             {"referenceId": bill_addr_id},
                    "saleOrderItems":             [],
                },
            }

        # ── Append line item ───────────────────────────────────────────────
        item_code     = _s(row, "Sale Order Item Code*")
        sku           = _s(row, "Item SKU Code*")
        shipping_meth = _s(row, "Shipping Method*") or "SHIPROCKET"
        gift_wrap     = _s(row, "Gift Wrap", "FALSE").upper() in ("TRUE", "1", "YES")
        on_hold       = _s(row, "On Hold",   "FALSE").upper() in ("TRUE", "1", "YES")

        try:
            quantity   = int(float(row.get("Quantity")     or 1))
        except (ValueError, TypeError):
            quantity   = 1
        try:
            packet_num = int(float(row.get("Packet Number") or 1))
        except (ValueError, TypeError):
            packet_num = 1

        item = {
            "code":               item_code,
            "itemSku":            sku,
            "shippingMethodCode": shipping_meth,
            "facilityCode":       _s(row, "Facility Code"),
            "packetNumber":       packet_num,
            "giftWrap":           gift_wrap,
            "onHold":             on_hold,
            "quantity":           quantity,
            "totalPrice":         _f(row, "Selling Price"),
            "sellingPrice":       _f(row, "Selling Price"),
            "prepaidAmount":      _f(row, "Prepaid Amount"),
            "discount":           _f(row, "Discount"),
            "shippingCharges":    _f(row, "Shipping Charges"),
            "storeCredit":        _f(row, "Store Credit"),
            "giftWrapCharges":    _f(row, "Gift Wrap Charges"),
        }

        voucher = _s(row, "Voucher Code")
        if voucher:
            item["voucherCode"]  = voucher
            item["voucherValue"] = _f(row, "Voucher Value")

        tracking = _s(row, "Tracking Number")
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

    # De-duplicate headers (handles the two "Date" columns in this sheet)
    seen    = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    # Find status column index upfront so we can read all statuses in one pass
    status_col = _find_or_create_status_col(ws, all_values)

    # Build per-row dicts, track earliest sheet row per order, cache statuses
    records:      list[dict]     = []
    row_index:    dict[str, int] = {}
    status_cache: dict[str, str] = {}

    for sheet_row, row_values in enumerate(all_values[1:], start=2):
        padded = row_values + [""] * (len(headers) - len(row_values))
        row    = dict(zip(headers, padded))
        records.append(row)

        order_code = str(row.get("Sales Order Code*", "")).strip()
        if not order_code:
            continue

        if order_code not in row_index:
            row_index[order_code]    = sheet_row
            # Status is in the same row_values if the column already exists
            status_val = padded[status_col - 1] if status_col <= len(padded) else ""
            status_cache[order_code] = str(status_val).strip()

    # Group into orders (only new ones will be processed below)
    orders = _group_rows_into_orders(records)

    url     = f"{TENANT_URL}/services/rest/v1/oms/saleOrder/create"
    results = {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    for order_code, order_data in orders.items():
        sheet_row      = row_index.get(order_code)
        current_status = status_cache.get(order_code, "")

        # Skip already processed orders
        if current_status in ("SUCCESS", "FAILED"):
            results["skipped"] += 1
            continue

        facility = order_data.pop("_facility_code", "")

        if not facility:
            status = "SKIPPED - missing Facility Code"
            if sheet_row:
                ws.update_cell(sheet_row, status_col, status)
            results["skipped"] += 1
            continue

        if not order_data["saleOrder"]["saleOrderItems"]:
            status = "SKIPPED - no items"
            if sheet_row:
                ws.update_cell(sheet_row, status_col, status)
            results["skipped"] += 1
            continue

        try:
            data = api_post(url, order_data, facility=facility)

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