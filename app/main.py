from fastapi import FastAPI
from fastapi.responses import JSONResponse

from app.routers import inventory, po

app = FastAPI(
    title="Uniware Automations API",
    description="REST API wrapper for Unicommerce/Uniware — inventory fetch and purchase order upload.",
    version="1.0.0",
)

app.include_router(inventory.router)
app.include_router(po.router)


@app.get("/health", tags=["Health"])
def health():
    return JSONResponse(content={"status": "ok"})

@app.get("/myip")
def my_ip():
    import requests
    r = requests.get("https://api.ipify.org?format=json", timeout=10)
    return r.json()
