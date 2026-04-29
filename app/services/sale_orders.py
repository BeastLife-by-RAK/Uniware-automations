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
    return str(row.get(key) or default).strip()


def _f(row: dict, key: str) -> float:
    try:
        return float(row.get(key) or 0)
    except (ValueError, TypeError):
        return 0.0


def _i(row: dict, key: str, default: int = 0) -> int:
    try:
        return int(float(row.get(key) or default))
    except (ValueError, TypeError):
        return default


def _b(row: dict, key: str, default: bool = False) -> bool:
    val = _s(row, key).upper()
    if val in ("TRUE", "1", "YES"):
        return True
    if val in ("FALSE", "0", "NO"):
        return False
    return default


def _opt(d: dict, key: str, value) -> None:
    """Set key on dict only when value is non-empty / non-None."""
    if value is not None and value != "":
        d[key] = value


def _parse_date(raw: str) -> str | None:
    if not raw:
        return None
    raw = raw.strip()
    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except ValueError:
            continue
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── GROUP ROWS INTO ORDERS ────────────────────────────────────────────────────
def _group_rows_into_orders(records: list[dict]) -> dict[str, dict]:
    orders: dict[str, dict] = {}

    for raw_row in records:
        row = {k: str(v).strip() if v is not None else "" for k, v in raw_row.items()}

        order_code = row.get("Sales Order Code*", "")
        if not order_code:
            continue

        # ── Build order header on first encounter ────────────────────────────
        if order_code not in orders:
            # Address ID always "1"; billing always mirrors shipping
            addr_id = "1"

            ship_addr: dict = {
                "id":           addr_id,
                "name":         _s(row, "Shipping Address Name"),
                "addressLine1": _s(row, "Shipping Address Line 1"),
                "city":         _s(row, "Shipping Address City"),
                "state":        _s(row, "Shipping Address State"),
                "country":      _s(row, "Shipping Address Country") or "India",
                "pincode":      _s(row, "Shipping Address Pincode"),
                "phone":        _s(row, "Shipping Address Phone"),
            }
            _opt(ship_addr, "addressLine2", _s(row, "Shipping Address Line 2"))
            _opt(ship_addr, "latitude",     _s(row, "Shipping Address Latitude"))
            _opt(ship_addr, "longitude",    _s(row, "Shipping Address Longitude"))

            # Billing mirrors shipping — single address entry, same referenceId
            addresses = [ship_addr]

            sale_order: dict = {
                "code":                       order_code,
                "displayOrderCode":           _s(row, "Display Sales Order Code") or order_code,
                "channel":                    "INFLUENCERS_MARKETING",
                "cashOnDelivery":             False,
                "customerName":               ship_addr["name"],
                "notificationMobile":         _s(row, "Notification Mobile") or ship_addr["phone"],
                "currencyCode":               _s(row, "Currency Code") or "INR",
                "totalPrepaidAmount":         _f(row, "Prepaid Amount"),
                "totalCashOnDeliveryCharges": 0,
                "totalDiscount":              _f(row, "Discount"),
                "totalShippingCharges":       _f(row, "Order Total Shipping Charges"),
                "totalGiftWrapCharges":       _f(row, "Gift Wrap Charges"),
                "totalStoreCredit":           _f(row, "Store Credit"),
                "addresses":                  addresses,
                "shippingAddress":            {"referenceId": addr_id},
                "billingAddress":             {"referenceId": addr_id},
                "saleOrderItems":             [],
            }

            # Optional order-level fields
            _opt(sale_order, "displayOrderDateTime",  _parse_date(_s(row, "Order Date as dd/mm/yyyy hh:MM:ss")))
            _opt(sale_order, "channelProcessingTime", _parse_date(
                _s(row, "Channel Order Processing Date as dd/MM/yyyy hh:mm:ss")
                or _s(row, "Channel Order Processing Date as dd/MM/yyyy hh:mm:ss_1")
            ))
            _opt(sale_order, "fulfillmentTat",          _parse_date(_s(row, "Fulfillment Tat")))
            _opt(sale_order, "customerCode",            _s(row, "Customer Code"))
            _opt(sale_order, "customerGSTIN",           _s(row, "Customer GSTIN"))
            _opt(sale_order, "priority",                _i(row, "Priority") or None)
            _opt(sale_order, "shippingPackageTypeCode", _s(row, "Shipping Package Type Code"))
            _opt(sale_order, "parentSaleOrderCode",     _s(row, "Parent Sale Order Code"))

            # shippingProviders — always Shiprocket1, no packetNumber at order level
            tracking = _s(row, "Tracking Number")
            sale_order["shippingProviders"] = [{
                "code":           "Shiprocket1",
                "packetNumber":   1,
                "trackingNumber": tracking,
            }]

            # saleOrderItemCombinations — only when both fields present
            combo_id   = _s(row, "Combination Identifier")
            combo_desc = _s(row, "Combination Description")
            if combo_id and combo_desc:
                sale_order["saleOrderItemCombinations"] = [{
                    "combinationIdentifier":  combo_id,
                    "combinationDescription": combo_desc,
                }]

            orders[order_code] = {
                "_facility_code": _s(row, "Facility Code"),
                "saleOrder": sale_order,
            }

        # ── Append line item ─────────────────────────────────────────────────
        sku      = _s(row, "Item SKU Code*")
        quantity = _i(row, "Quantity", 1)

        item: dict = {
            "code":               _s(row, "Sale Order Item Code*"),
            "itemSku":            sku,
            "shippingMethodCode": "STD",
            "facilityCode":       _s(row, "Facility Code"),
            "channelProductId":   sku,
            "packetNumber":       quantity,   # packet number = quantity
            "giftWrap":           _b(row, "Gift Wrap"),
            "quantity":           quantity,
            "totalPrice":         _f(row, "Selling Price"),
            "sellingPrice":       _f(row, "Selling Price"),
            "discount":           _f(row, "Discount"),
            "shippingCharges":    _f(row, "Shipping Charges"),
            "storeCredit":        _f(row, "Store Credit"),
            "giftWrapCharges":    _f(row, "Gift Wrap Charges"),
        }

        _opt(item, "giftMessage", _s(row, "Gift Message"))
        _opt(item, "voucherCode", _s(row, "Voucher Code"))
        if _s(row, "Voucher Code"):
            item["voucherValue"] = _f(row, "Voucher Value")

        orders[order_code]["saleOrder"]["saleOrderItems"].append(item)

    return orders


# ── MAIN ENTRY POINT ──────────────────────────────────────────────────────────
def process_sale_orders() -> dict:
    ws         = _get_sheet()
    all_values = ws.get_all_values()

    if not all_values or len(all_values) < 2:
        return {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    raw_headers = all_values[0]

    # De-duplicate headers (handles repeated column names e.g. two date cols)
    seen    = {}
    headers = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    records = []
    for row_values in all_values[1:]:
        padded = row_values + [""] * (len(headers) - len(row_values))
        records.append(dict(zip(headers, padded)))

    orders  = _group_rows_into_orders(records)
    url     = f"{TENANT_URL}/services/rest/v1/oms/saleOrder/create"
    results = {"success": 0, "failed": 0, "skipped": 0, "errors": []}

    for order_code, order_data in orders.items():
        facility = order_data.pop("_facility_code", "")

        if not facility:
            results["skipped"] += 1
            results["errors"].append({"order_id": order_code, "error": "Missing Facility Code"})
            continue

        if not order_data["saleOrder"]["saleOrderItems"]:
            results["skipped"] += 1
            results["errors"].append({"order_id": order_code, "error": "No items"})
            continue

        try:
            data = api_post(url, order_data, facility=facility)

            if data.get("successful"):
                uc_code = data.get("saleOrderDetailDTO", {}).get("code", order_code)
                results["success"] += 1
                print(f"  ✔ Order {order_code} → {uc_code}")
            else:
                errors = data.get("errors", [])
                msg    = errors[0].get("description") if errors else data.get("message", "Unknown error")
                results["failed"] += 1
                results["errors"].append({"order_id": order_code, "error": msg})
                print(f"  ✗ Order {order_code} failed: {msg}")

        except Exception as e:
            results["failed"] += 1
            results["errors"].append({"order_id": order_code, "error": str(e)})
            print(f"  ✗ Order {order_code} exception: {e}")

    return results