"""
Aquaponics dashboard backend.

Uses a persistent headless LibreOffice instance (via UNO) to recalculate the
REAL workbook on each request. This guarantees the dashboard's numbers match
the spreadsheet exactly — no formula reconstruction.

Endpoints
---------
GET  /health    liveness
GET  /options   valid dropdown values (greens, fish, regions) + current defaults
POST /calculate set core inputs, recalc, return NPV/verdict/charts

Concurrency note: a single LibreOffice document is not thread-safe, so
/calculate is serialized behind a lock. Recalcs are fast (well under a second),
so this is fine for expected traffic.
"""

import os, re, time, socket, subprocess, shutil, threading, atexit
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

# ── Config ──
SRC_XLSX = os.environ.get("MODEL_XLSX", "/app/Model_solver_multispecies_Parameters.xlsx")
WORK_XLSX = "/tmp/_model_work.xlsx"
UNO_PORT = 2002
HOME_DIR = os.environ.get("HOME", "/tmp")

GREENS = ["Tomato","Lettuce","Chicória","Almeirão Pão de Açúcar","Almeirão Amargo",
          "Espinafre","Alface","Salsinha","Manjericão","Agrião","Cebolinha"]
FISH = ["Tilapia","Trout"]
REGIONS = ["Africa","Asia","North America","South America","Europe","Oceania"]

DEFAULTS = dict(chosen_green="Lettuce", chosen_fish="Tilapia", total_area=1000,
                equity=10000, cost_of_equity=0.171, debt_interest_rate=0.1186,
                region="South America")

_lock = threading.Lock()
_doc = None
_soffice_proc = None


def _a1(sheet, a1):
    m = re.match(r'([A-Z]+)(\d+)', a1)
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return sheet.getCellByPosition(col - 1, int(m.group(2)) - 1)


def _start_soffice():
    global _soffice_proc
    env = {**os.environ, "HOME": HOME_DIR}
    subprocess.run(["pkill", "-f", "soffice"], env=env)
    time.sleep(2)
    _soffice_proc = subprocess.Popen(
        ["soffice","--headless","--invisible","--nocrashreport","--nodefault",
         "--norestore","--nologo",
         f"--accept=socket,host=localhost,port={UNO_PORT};urp;"],
        env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(60):
        try:
            s = socket.create_connection(("localhost", UNO_PORT), timeout=1); s.close(); return
        except OSError:
            time.sleep(1)
    raise RuntimeError("LibreOffice UNO port never opened")


def _open_doc():
    global _doc
    import uno
    from com.sun.star.beans import PropertyValue
    shutil.copy(SRC_XLSX, WORK_XLSX)
    lc = uno.getComponentContext()
    resolver = lc.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", lc)
    ctx = resolver.resolve(
        f"uno:socket,host=localhost,port={UNO_PORT};urp;StarOffice.ComponentContext")
    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    p = PropertyValue(); p.Name = "Hidden"; p.Value = True
    _doc = desktop.loadComponentFromURL("file://" + WORK_XLSX, "_blank", 0, (p,))
    _apply_ebitda_fix(_doc)


def _col_letter(idx0):
    """0-based column index -> A1 letters."""
    idx = idx0 + 1
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _apply_ebitda_fix(doc):
    """
    Fix a bug in the Framework P&L: the EBITDA row (T10:AC10) originally computed
    `= NetRevenue - SUM(COGS:People)`, which SKIPS the Water (row 8) and Energy
    (row 9) rows -- so climate-driven costs were displayed but never subtracted
    from cash flow. We rewrite each year's EBITDA to also subtract water & energy,
    so region/climate changes now correctly flow through to NPV.

    Columns T..AC are the 10 forecast years (col index 19..28), row 10 (index 9).
    Original: =<col>5 - SUM(<col>6:<col>7)
    Fixed:    =<col>5 - SUM(<col>6:<col>9)
    """
    fw = doc.Sheets.getByName("Framework")
    for col in range(19, 29):  # T..AC
        cell = fw.getCellByPosition(col, 9)  # row 10
        cl = _col_letter(col)
        cell.setFormula(f"={cl}5-SUM({cl}6:{cl}9)")


def _ensure_ready():
    global _doc
    if _doc is None:
        _start_soffice()
        _open_doc()


def _recalc(inp: dict) -> dict:
    _ensure_ready()
    fw = _doc.Sheets.getByName("Framework")
    fn = _doc.Sheets.getByName("Funding")
    inv = _doc.Sheets.getByName("Investment")

    # write inputs
    _a1(fw, "M5").setValue(float(inp["equity"]))
    _a1(fw, "M6").setValue(float(inp["total_area"]))
    _a1(fw, "M9").setString(inp["chosen_green"])
    _a1(fw, "M10").setString(inp["chosen_fish"])
    _a1(fw, "M11").setString(inp["region"])
    _a1(fn, "D5").setValue(float(inp["cost_of_equity"]))
    _a1(fn, "D8").setValue(float(inp["debt_interest_rate"]))

    _doc.calculateAll()

    npv = _a1(fw, "S20").getValue()
    verdict = _a1(fw, "T20").getString()
    payback = _a1(fw, "S22").getValue()
    wacc = _a1(fn, "L5").getValue()
    total_inv = _a1(inv, "K4").getValue()
    gross_rev = fw.getCellByPosition(20, 2).getValue()  # U3 year-2 gross revenue steady

    fcf = [round(fw.getCellByPosition(18+i, 15).getValue(), 2) for i in range(11)]
    acc = [round(fw.getCellByPosition(18+i, 17).getValue(), 2) for i in range(11)]
    wf_labels = ["Gross Revenue","Sales Taxes","COGS","People Costs","Water Costs",
                 "Energy Costs","Debt Interest","Corporate Tax","Investment",
                 "Loan Repayment","Free Cash Flow"]
    wf_vals = [round(fw.getCellByPosition(31, i).getValue(), 2) for i in range(2, 13)]
    waterfall = dict(zip(wf_labels, wf_vals))

    return dict(
        npv=round(npv, 2), verdict=verdict,
        payback_years=round(payback, 2) if payback and payback > 0 else None,
        wacc=round(wacc, 4), total_investment=round(total_inv, 2),
        gross_revenue=round(gross_rev, 2),
        free_cashflows=fcf, accumulated_value=acc, waterfall=waterfall,
    )


@asynccontextmanager
async def lifespan(app: FastAPI):
    with _lock:
        _ensure_ready()
    yield
    if _soffice_proc:
        _soffice_proc.terminate()

atexit.register(lambda: _soffice_proc.terminate() if _soffice_proc else None)

app = FastAPI(title="Aquaponics Dashboard", version="1.0.0", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


class CalcRequest(BaseModel):
    chosen_green: str = Field("Lettuce")
    chosen_fish: str = Field("Tilapia")
    total_area: float = Field(1000, ge=1)
    equity: float = Field(10000, ge=0)
    cost_of_equity: float = Field(0.171, ge=0, le=2)
    debt_interest_rate: float = Field(0.1186, ge=0, le=2)
    region: str = Field("South America")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/options")
def options():
    return dict(greens=GREENS, fish=FISH, regions=REGIONS, defaults=DEFAULTS)


@app.post("/calculate")
def calculate(req: CalcRequest):
    if req.chosen_green not in GREENS:
        raise HTTPException(400, f"Unknown green: {req.chosen_green}")
    if req.chosen_fish not in FISH:
        raise HTTPException(400, f"Unknown fish: {req.chosen_fish}")
    if req.region not in REGIONS:
        raise HTTPException(400, f"Unknown region: {req.region}")
    try:
        with _lock:
            return _recalc(req.model_dump())
    except Exception as e:
        # recover a dead LibreOffice on next call
        global _doc
        _doc = None
        raise HTTPException(500, f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("backend:app", host="0.0.0.0", port=8000)
