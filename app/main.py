from fastapi import FastAPI
from fastapi.responses import JSONResponse

from app.routers import inventory, po, sale_orders

app = FastAPI(
    title="Uniware Automations API",
    description="REST API wrapper for Unicommerce/Uniware — inventory fetch, PO upload and sale order creation.",
    version="1.0.0",
)

app.include_router(inventory.router)
app.include_router(po.router)
app.include_router(sale_orders.router)


@app.get("/health", tags=["Health"])
def health():
    return JSONResponse(content={"status": "ok"})