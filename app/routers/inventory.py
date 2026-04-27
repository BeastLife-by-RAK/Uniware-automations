import io
from typing import Optional
from fastapi import APIRouter, Query, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse

from app.services.inventory import fetch_inventory, build_inventory_excel

router = APIRouter(prefix="/inventory", tags=["Inventory"])


@router.get("/fetch")
def get_inventory(
    format: str = Query("json", description="Response format: json or excel"),
    updated_since_minutes: Optional[int] = Query(None, description="Only SKUs updated in last N minutes (max 1440)"),
    skus: Optional[str] = Query(None, description="Comma-separated SKU codes e.g. SKU001,SKU002"),
):
    """
    Fetch inventory snapshot from Unicommerce.

    - **format=json** → raw JSON (default, good for Google Sheets)
    - **format=excel** → downloadable .xlsx file
    - **updated_since_minutes** → only recently changed SKUs
    - **skus** → filter to specific SKUs
    """
    sku_list = [s.strip() for s in skus.split(",")] if skus else None

    try:
        records = fetch_inventory(updated_since_minutes=updated_since_minutes, sku_list=sku_list)
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e))

    if format == "excel":
        return StreamingResponse(
            io.BytesIO(build_inventory_excel(records)),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=inventory_snapshot.xlsx"},
        )

    return JSONResponse(content={"count": len(records), "records": records})
