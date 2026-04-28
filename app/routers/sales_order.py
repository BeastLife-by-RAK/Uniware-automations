from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse

from app.services.sale_orders import process_sale_orders

router = APIRouter(prefix="/orders", tags=["Sale Orders"])


@router.post("/process-influencer")
def process_influencer_orders():
    """
    Read influencer sale orders from Google Sheet and create them in Unicommerce.
    Skips rows already marked SUCCESS or FAILED.
    Writes Order Status back to the sheet.
    """
    try:
        results = process_sale_orders()
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e))

    return JSONResponse(content={
        "message":  "Influencer orders processed",
        "success":  results["success"],
        "failed":   results["failed"],
        "skipped":  results["skipped"],
        "errors":   results["errors"],
    })