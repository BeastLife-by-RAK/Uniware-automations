import io
from typing import Optional
from fastapi import APIRouter, Query, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse

from app.services.inventory import fetch_inventory, build_inventory_excel, fetch_facilities
from app.services.sheets import push_inventory_to_sheets

router = APIRouter(prefix="/inventory", tags=["Inventory"])


@router.get("/facilities")
def list_facilities():
    """List all facility codes."""
    try:
        codes = fetch_facilities()
        return JSONResponse(content={"count": len(codes), "facilities": codes})
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e))


@router.get("/fetch")
def get_inventory(
    format: str = Query("json", description="Response format: json or excel"),
    updated_since_minutes: Optional[int] = Query(
        None,
        description="Only SKUs updated in last N minutes. Omit for full snapshot."
    ),
    skus: Optional[str] = Query(
        None,
        description="Comma-separated SKU codes e.g. SKU001,SKU002"
    ),
):
    """
    Fetch inventory snapshot from Unicommerce across all facilities.

    - **format=json** → raw JSON (default)
    - **format=excel** → downloadable .xlsx file
    - **updated_since_minutes** → SKUs updated in last N minutes. Omit for full snapshot.
    - **skus** → filter to specific SKUs
    """
    sku_list = [s.strip() for s in skus.split(",")] if skus else None

    try:
        records = fetch_inventory(
            updated_since_minutes=updated_since_minutes,
            sku_list=sku_list,
        )
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e))

    if format == "excel":
        return StreamingResponse(
            io.BytesIO(build_inventory_excel(records)),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=inventory_snapshot.xlsx"},
        )

    return JSONResponse(content={"count": len(records), "records": records})


@router.post("/push-to-sheets")
def push_to_sheets(
    updated_since_minutes: Optional[int] = Query(
        None,
        description="Only SKUs updated in last N minutes. Omit for full snapshot."
    ),
):
    """
    Fetch inventory from Unicommerce and push directly to Google Sheets.
    Creates a new dated tab per facility e.g. GGN-2026-04-27.
    Designed to be triggered by Cloud Scheduler.
    """
    try:
        records = fetch_inventory(updated_since_minutes=updated_since_minutes)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Unicommerce fetch failed: {e}")

    if not records:
        return JSONResponse(content={"message": "No records returned", "tabs_created": {}})

    try:
        summary = push_inventory_to_sheets(records)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Google Sheets write failed: {e}")

    return JSONResponse(content={
        "message": "Successfully pushed to Google Sheets",
        "tabs_created": summary,
        "total_records": len(records),
    })

#additional function for testing 
@router.get("/audit")
def audit_missing_skus():
    """
    Compare catalog SKUs vs inventory snapshot SKUs.
    Returns which SKUs are missing from the inventory response.
    """
    from app.services.auth import api_post, TENANT_URL

    # Step 1 — get ALL SKUs from catalog
    search_url = f"{TENANT_URL}/services/rest/v1/product/itemType/search"
    catalog_skus = set()
    start = 0

    while True:
        payload = {
            "searchOptions": {
                "displayLength": 500,
                "displayStart": start,
                "getCount": True
            }
        }
        data = api_post(search_url, payload)
        elements = data.get("elements", [])
        for item in elements:
            if item.get("skuCode") and item.get("enabled", True):
                catalog_skus.add(item["skuCode"])
        start += len(elements)
        if not elements or start >= data.get("totalRecords", 0):
            break

    # Step 2 — get SKUs from inventory snapshot (current method)
    inv_url = f"{TENANT_URL}/services/rest/v1/inventory/inventorySnapshot/get"
    inventory_skus = set()

    for code in fetch_facilities():
        try:
            data = api_post(inv_url, {"updatedSinceInMinutes": 1440}, facility=code)
            for r in data.get("inventorySnapshots") or data.get("inventorySnapShotList") or []:
                inventory_skus.add(r.get("itemTypeSKU"))
        except:
            pass

    # Step 3 — compare
    missing = sorted(catalog_skus - inventory_skus)

    return JSONResponse(content={
        "catalog_total":    len(catalog_skus),
        "inventory_total":  len(inventory_skus),
        "missing_count":    len(missing),
        "missing_skus":     missing,
    })
