import io
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from app.services.auth import api_post, api_get, TENANT_URL, FACILITY_CODE


def fetch_facilities() -> list[str]:
    return [
        "Emiza_B2C_BLR",
        "Emiza_B2C_GGN",
        "Emiza_B2C_Mumbai",
        "Emiza_B2C_WB",
    ]


def fetch_inventory(
    updated_since_minutes: Optional[int] = None,
    sku_list: Optional[list[str]] = None,
) -> list[dict]:
    url            = f"{TENANT_URL}/services/rest/v1/inventory/inventorySnapshot/get"
    facility_codes = fetch_facilities()

    # Default to last 24 hours if no filter provided
    if not updated_since_minutes and not sku_list:
        updated_since_minutes = 1440

    all_records = []
    seen_keys   = set()

    for code in facility_codes:
        payload = {}
        if sku_list:
            payload["itemTypeSKUs"] = sku_list
        if updated_since_minutes:
            payload["updatedSinceInMinutes"] = updated_since_minutes

        try:
            data = api_post(url, payload, facility=code)

            print(f"  Raw response for {code}: {str(data)[:500]}")

            if data.get("successful"):
                records = (
                    data.get("inventorySnapshots")
                    or data.get("inventorySnapShotList")
                    or []
                )
                for r in records:
                    key = f"{r.get('itemTypeSKU')}_{code}"
                    if key not in seen_keys:
                        seen_keys.add(key)
                        r["facilityCode"] = code
                        all_records.append(r)
                print(f"  ✓ Facility {code}: {len(records)} SKUs")
            else:
                print(f"  ⚠ Facility {code}: {data.get('message')} | errors: {data.get('errors')}")

        except Exception as e:
            print(f"  ⚠ Facility {code} failed: {e}")
            continue

    print(f"Total records across all facilities: {len(all_records)}")
    return all_records


def build_inventory_excel(records: list[dict]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventory Snapshot"

    header_font  = Font(bold=True, color="FFFFFF", name="Arial")
    header_fill  = PatternFill("solid", start_color="1F4E79")
    center_align = Alignment(horizontal="center", vertical="center")

    headers = [
        "SKU Code", "Item Name", "Facility Code",
        "Sellable Qty", "Blocked Qty", "Pending Putaway",
        "Bad Inventory", "Virtual Inventory", "Updated At",
    ]
    for col, h in enumerate(headers, start=1):
        cell           = ws.cell(row=1, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align
    ws.row_dimensions[1].height = 20

    for row_idx, item in enumerate(records, start=2):
        ws.cell(row=row_idx, column=1, value=item.get("itemTypeSKU"))
        ws.cell(row=row_idx, column=2, value=item.get("itemName"))
        ws.cell(row=row_idx, column=3, value=item.get("facilityCode"))
        ws.cell(row=row_idx, column=4, value=item.get("inventory", 0))
        ws.cell(row=row_idx, column=5, value=item.get("blockedInventory", 0))
        ws.cell(row=row_idx, column=6, value=item.get("pendingPutawayInventory", 0))
        ws.cell(row=row_idx, column=7, value=item.get("badInventory", 0))
        ws.cell(row=row_idx, column=8, value=item.get("virtualInventory", 0))
        ws.cell(row=row_idx, column=9, value=item.get("updatedAt"))

        fill_color = "EBF3FB" if row_idx % 2 == 0 else "FFFFFF"
        row_fill   = PatternFill("solid", start_color=fill_color)
        for col in range(1, 10):
            ws.cell(row=row_idx, column=col).fill = row_fill

    col_widths = [18, 30, 16, 14, 14, 16, 16, 18, 22]
    for col, width in enumerate(col_widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width

    last = len(records) + 2
    ws.cell(row=last, column=1, value="TOTAL").font = Font(bold=True, name="Arial")
    ws.cell(row=last, column=4, value=f"=SUM(D2:D{last-1})")
    ws.cell(row=last, column=5, value=f"=SUM(E2:E{last-1})")
    total_fill = PatternFill("solid", start_color="D6E4F0")
    for col in range(1, 10):
        ws.cell(row=last, column=col).fill = total_fill

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()