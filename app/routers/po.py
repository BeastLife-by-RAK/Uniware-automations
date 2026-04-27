import io
from fastapi import APIRouter, UploadFile, File, Query, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse

from app.services.po import parse_po_excel, create_purchase_orders, build_results_excel, build_po_template

router = APIRouter(prefix="/po", tags=["Purchase Orders"])


@router.get("/template")
def download_po_template():
    """Download a blank Excel template with PO Header and PO Items sheets."""
    return StreamingResponse(
        io.BytesIO(build_po_template()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=po_template.xlsx"},
    )


@router.post("/upload")
async def upload_purchase_orders(
    file: UploadFile = File(..., description="Filled PO Excel (PO Header + PO Items sheets)"),
    approve: bool    = Query(False, description="Create POs in APPROVED status directly"),
    dry_run: bool    = Query(False, description="Parse only — no API call to Unicommerce"),
):
    """
    Upload a filled PO Excel to create Purchase Orders in Unicommerce.

    Returns a downloadable Excel with per-PO SUCCESS / FAILED / SKIPPED results.
    Use **dry_run=true** to validate your Excel before submitting.
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Only .xlsx or .xls files accepted")

    file_bytes = await file.read()

    try:
        po_list, skipped = parse_po_excel(file_bytes)
    except KeyError as e:
        raise HTTPException(status_code=400, detail=f"Missing sheet or column: {e}")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to parse Excel: {e}")

    if not po_list:
        raise HTTPException(status_code=400, detail="No valid POs found in the uploaded file")

    if dry_run:
        return JSONResponse(content={
            "dry_run":  True,
            "po_count": len(po_list),
            "skipped":  skipped,
            "payloads": po_list,
        })

    try:
        results = create_purchase_orders(po_list, approve=approve)
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e))

    success = sum(1 for r in results if r["status"] == "SUCCESS")
    failed  = sum(1 for r in results if r["status"] in ("FAILED", "ERROR"))

    return StreamingResponse(
        io.BytesIO(build_results_excel(results)),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=po_upload_results.xlsx",
            "X-Total":   str(len(results)),
            "X-Success": str(success),
            "X-Failed":  str(failed),
        },
    )
