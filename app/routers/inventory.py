#OLD CODE 
# import io
# from typing import Optional
# from fastapi import APIRouter, Query, HTTPException
# from fastapi.responses import StreamingResponse, JSONResponse

# from app.services.inventory import fetch_inventory, build_inventory_excel, fetch_facilities
# from app.services.sheets import push_inventory_to_sheets

# router = APIRouter(prefix="/inventory", tags=["Inventory"])


# @router.get("/facilities")
# def list_facilities():
#     """List all facility codes."""
#     try:
#         codes = fetch_facilities()
#         return JSONResponse(content={"count": len(codes), "facilities": codes})
#     except Exception as e:
#         raise HTTPException(status_code=502, detail=str(e))


# @router.get("/fetch")
# def get_inventory(
#     format: str = Query("json", description="Response format: json or excel"),
#     updated_since_minutes: Optional[int] = Query(
#         None,
#         description="Only SKUs updated in last N minutes. Omit for full snapshot."
#     ),
#     skus: Optional[str] = Query(
#         None,
#         description="Comma-separated SKU codes e.g. SKU001,SKU002"
#     ),
# ):
#     """
#     Fetch inventory snapshot from Unicommerce across all facilities.

#     - **format=json** → raw JSON (default)
#     - **format=excel** → downloadable .xlsx file
#     - **updated_since_minutes** → SKUs updated in last N minutes. Omit for full snapshot.
#     - **skus** → filter to specific SKUs
#     """
#     sku_list = [s.strip() for s in skus.split(",")] if skus else None

#     try:
#         records = fetch_inventory(
#             updated_since_minutes=updated_since_minutes,
#             sku_list=sku_list,
#         )
#     except Exception as e:
#         raise HTTPException(status_code=502, detail=str(e))

#     if format == "excel":
#         return StreamingResponse(
#             io.BytesIO(build_inventory_excel(records)),
#             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             headers={"Content-Disposition": "attachment; filename=inventory_snapshot.xlsx"},
#         )

#     return JSONResponse(content={"count": len(records), "records": records})


# @router.post("/push-to-sheets")
# def push_to_sheets(
#     updated_since_minutes: Optional[int] = Query(
#         None,
#         description="Only SKUs updated in last N minutes. Omit for full snapshot."
#     ),
# ):
#     """
#     Fetch inventory from Unicommerce and push directly to Google Sheets.
#     Creates a new dated tab per facility e.g. GGN-2026-04-27.
#     Designed to be triggered by Cloud Scheduler.
#     """
#     try:
#         records = fetch_inventory(updated_since_minutes=updated_since_minutes)
#     except Exception as e:
#         raise HTTPException(status_code=502, detail=f"Unicommerce fetch failed: {e}")

#     if not records:
#         return JSONResponse(content={"message": "No records returned", "tabs_created": {}})

#     try:
#         summary = push_inventory_to_sheets(records)
#     except Exception as e:
#         raise HTTPException(status_code=502, detail=f"Google Sheets write failed: {e}")

#     return JSONResponse(content={
#         "message": "Successfully pushed to Google Sheets",
#         "tabs_created": summary,
#         "total_records": len(records),
#     })

# #additional function for testing 
# @router.get("/audit")
# def audit_missing_skus():
#     """
#     Compare catalog SKUs vs inventory snapshot SKUs.
#     Returns which SKUs are missing from the inventory response.
#     """
#     from app.services.auth import api_post, TENANT_URL

#     # Step 1 — get ALL SKUs from catalog
#     search_url = f"{TENANT_URL}/services/rest/v1/product/itemType/search"
#     catalog_skus = set()
#     start = 0

#     while True:
#         payload = {
#             "searchOptions": {
#                 "displayLength": 500,
#                 "displayStart": start,
#                 "getCount": True
#             }
#         }
#         data = api_post(search_url, payload)
#         elements = data.get("elements", [])
#         for item in elements:
#             if item.get("skuCode") and item.get("enabled", True):
#                 catalog_skus.add(item["skuCode"])
#         start += len(elements)
#         if not elements or start >= data.get("totalRecords", 0):
#             break

#     # Step 2 — get SKUs from inventory snapshot (current method)
#     inv_url = f"{TENANT_URL}/services/rest/v1/inventory/inventorySnapshot/get"
#     inventory_skus = set()

#     for code in fetch_facilities():
#         try:
#             data = api_post(inv_url, {"updatedSinceInMinutes": 1440}, facility=code)
#             for r in data.get("inventorySnapshots") or data.get("inventorySnapShotList") or []:
#                 inventory_skus.add(r.get("itemTypeSKU"))
#         except:
#             pass

#     # Step 3 — compare
#     missing = sorted(catalog_skus - inventory_skus)

#     return JSONResponse(content={
#         "catalog_total":    len(catalog_skus),
#         "inventory_total":  len(inventory_skus),
#         "missing_count":    len(missing),
#         "missing_skus":     missing,
#     })
#NEW CODE 
import io
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from app.services.auth import api_post, api_get, TENANT_URL, FACILITY_CODE


FACILITY_MAP = {
    "Emiza_B2C_BLR":    "Bangalore",
    "Emiza_B2C_GGN":    "Gurugram",
    "Emiza_B2C_Mumbai": "Mumbai",
    "Emiza_B2C_WB":     "West Bengal",
}

_SKU_PAGE_SIZE   = 500   # max safe page size for itemType/search
_INV_CHUNK_SIZE  = 500   # inventory snapshot accepts up to 10,000; keep chunks small


def fetch_facilities() -> list[str]:
    return list(FACILITY_MAP.keys())


def fetch_all_sku_codes() -> list[str]:
    """
    Paginate through /product/itemType/search to collect every enabled SKU code.
    Runs once per fetch_inventory call — replaces the updatedSinceInMinutes filter
    so SKUs at 0 stock (sold out, awaiting restock) are never silently dropped.
    """
    url   = f"{TENANT_URL}/services/rest/v1/product/itemType/search"
    start = 0
    skus  = []

    while True:
        payload = {
            "searchOptions": {
                "displayLength": _SKU_PAGE_SIZE,
                "displayStart":  start,
                "getCount":      True,
            }
        }
        data = api_post(url, payload)   # Tenant-level — no facility header needed

        if not data.get("successful"):
            print(f"  ⚠ SKU catalog fetch failed: {data.get('message')}")
            break

        elements = data.get("elements", [])
        for item in elements:
            sku = item.get("skuCode")
            if sku and item.get("enabled", True):
                skus.append(sku)

        start += len(elements)
        total  = data.get("totalRecords", 0)

        if not elements or start >= total:
            break

    print(f"  ✔ Catalog: {len(skus)} active SKUs found")
    return skus


def fetch_inventory(
    updated_since_minutes: Optional[int] = None,
    sku_list: Optional[list[str]] = None,
) -> list[dict]:
    url            = f"{TENANT_URL}/services/rest/v1/inventory/inventorySnapshot/get"
    facility_codes = fetch_facilities()

    # Resolve which SKUs to query:
    # 1. Explicit list passed by caller (e.g. from router ?skus= param)
    # 2. Full catalog fetch — guarantees sold-out SKUs are never dropped
    # updatedSinceInMinutes is intentionally NOT used as the sole filter anymore
    # because it silently excludes SKUs with no inventory activity in 24hrs.
    if sku_list:
        skus = sku_list
        print(f"  Using caller-supplied SKU list ({len(skus)} SKUs)")
    else:
        skus = fetch_all_sku_codes()

    all_records = []
    seen_keys   = set()

    for code in facility_codes:
        try:
            facility_records = []

            # Chunk SKUs — API hard limit is 10,000 but smaller chunks are safer
            for i in range(0, len(skus), _INV_CHUNK_SIZE):
                chunk   = skus[i : i + _INV_CHUNK_SIZE]
                payload = {"itemTypeSKUs": chunk}
                data    = api_post(url, payload, facility=code)

                if data.get("successful"):
                    records = (
                        data.get("inventorySnapshots")
                        or data.get("inventorySnapShotList")
                        or []
                    )
                    facility_records.extend(records)
                else:
                    print(f"  ⚠ {code} chunk {i // _INV_CHUNK_SIZE + 1}: {data.get('message')}")

            for r in facility_records:
                key = f"{r.get('itemTypeSKU')}_{code}"
                if key not in seen_keys:
                    seen_keys.add(key)
                    r["facilityCode"]     = code
                    r["facilityLocation"] = FACILITY_MAP.get(code, code)
                    all_records.append(r)

            print(f"  ✔ Facility {code} ({FACILITY_MAP.get(code)}): {len(facility_records)} SKUs")

        except Exception as e:
            print(f"  ⚠ Facility {code} failed: {e}")
            continue

    print(f"Total records across all facilities: {len(all_records)}")
    return all_records


def build_inventory_excel(records: list[dict]) -> bytes:
    # ... your existing build_inventory_excel function stays exactly the same
