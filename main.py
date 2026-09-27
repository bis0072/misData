import os
import datetime
from pathlib import Path
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from routers.api import router as api_router

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

app = FastAPI(
    title="MIS HTML to Excel Service",
    description="High-performance, zero-disk-storage HTML report converter with VLOOKUP and Excel styling.",
    version="1.0.0"
)

# Enable CORS for external API integrations
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static assets (CSS, JS)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Jinja2 template engine configuration
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

# Include modular API endpoints (/health, /api/convert, /api/summary)
app.include_router(api_router)


@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    """Serves the web portal rendered dynamically via Jinja2Templates."""
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "title": "MIS HTML to Excel Studio",
            "default_table_index": 4,
            "default_nop_col_index": 7,
            "default_prem_col_index": 9,
            "current_year": datetime.date.today().year
        }
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)
