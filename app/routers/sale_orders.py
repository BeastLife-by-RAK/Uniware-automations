#OLD CODE 
# from fastapi import APIRouter, HTTPException
# from fastapi.responses import JSONResponse

# from app.services.sale_orders import process_sale_orders

# router = APIRouter(prefix="/orders", tags=["Sale Orders"])


# @router.post("/process-influencer")
# def process_influencer_orders():
#     """
#     Read influencer sale orders from Google Sheet and create them in Unicommerce.
#     Skips rows already marked SUCCESS or FAILED.
#     Writes Order Status back to the sheet.
#     """
#     try:
#         results = process_sale_orders()
#     except Exception as e:
#         raise HTTPException(status_code=502, detail=str(e))

#     return JSONResponse(content={
#         "message":  "Influencer orders processed",
#         "success":  results["success"],
#         "failed":   results["failed"],
#         "skipped":  results["skipped"],
#         "errors":   results["errors"],
#     })

#NEW CODE
from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import JSONResponse

from app.services.sale_orders import process_sale_orders

router = APIRouter(prefix="/orders", tags=["Sale Orders"])


@router.post("/process-influencer")
def process_influencer_orders(
    dry_run: bool = Query(False, description="Preview payloads without submitting to Unicommerce"),
):
    """
    Read influencer sale orders from Google Sheet and create them in Unicommerce.
    Skips rows already marked SUCCESS.
    Writes Order Status back to the sheet.
    Use dry_run=true to preview payloads without submitting anything.
    """
    try:
        results = process_sale_orders(dry_run=dry_run)
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e))

    return JSONResponse(content={
        "message":  "DRY RUN — nothing submitted" if dry_run else "Influencer orders processed",
        "dry_run":  dry_run,
        "success":  results["success"],
        "failed":   results["failed"],
        "skipped":  results["skipped"],
        "errors":   results["errors"],
        "payloads": results.get("dry_run_payloads", []),
    })
