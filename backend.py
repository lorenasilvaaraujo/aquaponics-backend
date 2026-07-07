"""
Aquaponics dashboard backend — lightweight edition.

Uses the `formulas` library to evaluate the workbook's Excel formulas directly
in Python. No LibreOffice, so memory stays well under 512 MB and it fits
Render's free tier.

The compiled model is loaded ONCE at startup (a few seconds), then each
/calculate call runs the dependency graph with the user's input overrides
(~1 second).

Fidelity: verified to reproduce the workbook's NPV to the penny
(Framework!S20 = 5913.99) and to respond correctly to every core lever.

Two spreadsheet-level corrections are baked into the model file we ship
(Model_fixed.xlsx), NOT hacked in code:
  1. PARTS_PERCENTAGE named range (was missing -> broke Investment!K4 and the
     whole funding/NPV chain). Defined to Investment!$B$7, matching intent.
  2. EBITDA row (Framework T10:AC10) rewritten to subtract Water (row 8) and
     Energy (row 9), which the original formula skipped -- so region/climate
     now correctly flows into cash flow and NPV.
"""

import os
import threading
import warnings
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

warnings.filterwarnings("ignore")

MODEL_XLSX = os.environ.get("MODEL_XLSX", "/app/Model_fixed.xlsx")
MODEL_TAG = "'[" + os.path.basename(MODEL_XLSX) + "]FRAMEWORK'"
FUND_TAG = "'[" + os.path.basename(MODEL_XLSX) + "]FUNDING'"

GREENS = ["Tomato","Lettuce","Chicória","Almeirão Pão de Açúcar","Almeirão Amargo",
          "Espinafre","Alface","Salsinha","Manjericão","Agrião","Cebolinha"]
FISH = ["Tilapia","Trout"]
REGIONS = ["Africa","Asia","North America","South America","Europe","Oceania"]

DEFAULTS = dict(chosen_green="Lettuce", chosen_fish="Tilapia", total_area=1000,
                equity=10000, cost_of_equity=0.171, debt_interest_rate=0.1186,
                region="South America")

CELL = dict(
    chosen_green      = f"{MODEL_TAG}!M9",
    chosen_fish       = f"{MODEL_TAG}!M10",
    total_area        = f"{MODEL_TAG}!M6",
    equity            = f"{MODEL_TAG}!M5",
    region            = f"{MODEL_TAG}!M11",
    cost_of_equity    = f"{FUND_TAG}!D5",
    debt_interest_rate= f"{FUND_TAG}!D8",
)

_model = None
_lock = threading.Lock()


def _load_model():
    global _model
    import formulas
    _model = formulas.ExcelModel().loads(MODEL_XLSX).finish()


def _cell_value(sol, sheet, a1):
    key_suffix = "!" + a1
    for k in sol:
        ku = k.upper()
        if sheet in ku and ku.endswith(key_suffix):
            v = sol[k].value
            try:
                return v[0][0]
            except (TypeError, IndexError):
                return v
    return None


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def _calculate(inp: dict) -> dict:
    if _model is None:
        _load_model()

    overrides = {
        CELL["chosen_green"]: inp["chosen_green"],
        CELL["chosen_fish"]:  inp["chosen_fish"],
        CELL["total_area"]:   inp["total_area"],
        CELL["equity"]:       inp["equity"],
        CELL["region"]:       inp["region"],
        CELL["cost_of_equity"]:     inp["cost_of_equity"],
        CELL["debt_interest_rate"]: inp["debt_interest_rate"],
    }
    sol = _model.calculate(inputs=overrides)

    npv = _num(_cell_value(sol, "FRAMEWORK", "S20"))
    verdict = _cell_value(sol, "FRAMEWORK", "T20")
    if isinstance(verdict, str):
        verdict = verdict.strip()
    else:
        verdict = "YES" if (npv or 0) > 0 else "NO"
    payback = _num(_cell_value(sol, "FRAMEWORK", "S22"))
    wacc = _num(_cell_value(sol, "FUNDING", "L5"))
    total_inv = _num(_cell_value(sol, "INVESTMENT", "K4"))
    gross_rev = _num(_cell_value(sol, "FRAMEWORK", "U3"))

    cols = ["S","T","U","V","W","X","Y","Z","AA","AB","AC"]
    fcf = [_num(_cell_value(sol, "FRAMEWORK", c + "16")) or 0 for c in cols]
    acc = [_num(_cell_value(sol, "FRAMEWORK", c + "18")) or 0 for c in cols]

    wf_labels = ["Gross Revenue","Sales Taxes","COGS","People Costs","Water Costs",
                 "Energy Costs","Debt Interest","Corporate Tax","Investment",
                 "Loan Repayment","Free Cash Flow"]
    wf_vals = [_num(_cell_value(sol, "FRAMEWORK", f"AF{3+i}")) or 0 for i in range(11)]
    waterfall = dict(zip(wf_labels, wf_vals))

    return dict(
        npv=round(npv, 2) if npv is not None else None,
        verdict=verdict,
        payback_years=round(payback, 2) if payback and payback > 0 else None,
        wacc=round(wacc, 4) if wacc is not None else None,
        total_investment=round(total_inv, 2) if total_inv is not None else None,
        gross_revenue=round(gross_rev, 2) if gross_rev is not None else None,
        free_cashflows=[round(x, 2) for x in fcf],
        accumulated_value=[round(x, 2) for x in acc],
        waterfall={k: round(v, 2) for k, v in waterfall.items()},
    )


@asynccontextmanager
async def lifespan(app: FastAPI):
    try:
        with _lock:
            _load_model()
    except Exception as e:
        print(f"[startup] model load deferred: {e}", flush=True)
    yield


app = FastAPI(title="Aquaponics Dashboard", version="2.0.0", lifespan=lifespan)
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
        raise HTTPException(400, f"Unknown crop: {req.chosen_green}")
    if req.chosen_fish not in FISH:
        raise HTTPException(400, f"Unknown fish: {req.chosen_fish}")
    if req.region not in REGIONS:
        raise HTTPException(400, f"Unknown region: {req.region}")
    try:
        with _lock:
            return _calculate(req.model_dump())
    except Exception as e:
        raise HTTPException(500, f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("backend:app", host="0.0.0.0", port=8000)
