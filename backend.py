"""
Aquaponics dashboard backend — guided-flow edition.

Uses the `formulas` library to evaluate the workbook directly (no LibreOffice,
low memory). Loaded once at startup; each /calculate runs the dependency graph
with the user's overrides.

Flow supported:
  Step 1  Location  -> city's monthly climate (from the frontend, Open-Meteo)
                       + auto-derived region (energy/water price defaults)
  Step 2  Species   -> chosen crop + fish
  Step 3  Prices &  -> crop/fish sell price, crop younglings, substrate,
          costs        fish younglings, fish feed, energy price, water price

Spreadsheet corrections baked into Model_fixed.xlsx (not code):
  1. PARTS_PERCENTAGE named range (was missing) -> Investment!$B$7.
  2. EBITDA row (Framework T10:AC10) now subtracts Water & Energy.
"""

import os
import threading
import warnings
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import Optional, List

warnings.filterwarnings("ignore")

MODEL_XLSX = os.environ.get("MODEL_XLSX", "/app/Model_fixed.xlsx")
BASENAME = os.path.basename(MODEL_XLSX)

def _tag(sheet):
    return f"'[{BASENAME}]{sheet.upper()}'"

GREENS = ["Tomato","Lettuce","Chicória","Almeirão Pão de Açúcar","Almeirão Amargo",
          "Espinafre","Alface","Salsinha","Manjericão","Agrião","Cebolinha"]
FISH = ["Tilapia","Trout"]
REGIONS = ["Africa","Asia","North America","South America","Europe","Oceania"]

# Row of each species in the Revenues price table (A17:C31) and Costs tables.
GREEN_ROW_REV = {g: 17+i for i, g in enumerate(GREENS)}          # Revenues C17..C27
FISH_ROW_REV = {"Tilapia": 30, "Trout": 31}                      # Revenues C30..C31
GREEN_ROW_COST = {g: 42+i for i, g in enumerate(GREENS)}         # Costs E/F 42..52
FISH_ROW_COST = {"Tilapia": 63, "Trout": 64}                     # Costs E/F 63..64
REGION_ROW_PARAM = {r: 55+i for i, r in enumerate(REGIONS)}      # Parameters C/F 55..60

# Regional default energy (avg col C) and water (avg col F) — for pre-filling the UI.
REGION_DEFAULTS = {
    "Africa":        dict(energy=0.133, water=0.5),
    "Asia":          dict(energy=0.107, water=0.3),
    "North America": dict(energy=0.165, water=1.0),
    "South America": dict(energy=0.203, water=0.8),
    "Europe":        dict(energy=0.240, water=3.5),
    "Oceania":       dict(energy=0.266, water=4.0),
}

# Species default prices/costs (for pre-filling the UI on selection).
GREEN_PRICE_DEFAULT = {"Tomato":3,"Lettuce":5,"Chicória":10,"Almeirão Pão de Açúcar":10,
    "Almeirão Amargo":10,"Espinafre":15,"Alface":5,"Salsinha":15,"Manjericão":30,
    "Agrião":15,"Cebolinha":12}
FISH_PRICE_DEFAULT = {"Tilapia":10,"Trout":15}
# younglings (E), substrate (F) defaults — all non-tomato greens share these
GREEN_COST_DEFAULT = {g: dict(young=0.04, substrate=0.014) for g in GREENS}
GREEN_COST_DEFAULT["Tomato"] = dict(young=0.5, substrate=32.2)
FISH_COST_DEFAULT = {"Tilapia": dict(young=0.16, feed=2), "Trout": dict(young=0.16, feed=3)}

# ISO country code -> region. Covers common cases; default South America? No —
# we default unknown to the continent via a broad fallback map.
COUNTRY_REGION = {
    # Africa
    "DZ":"Africa","AO":"Africa","EG":"Africa","ET":"Africa","GH":"Africa","KE":"Africa",
    "MA":"Africa","NG":"Africa","ZA":"Africa","TZ":"Africa","TN":"Africa","UG":"Africa",
    "SN":"Africa","CI":"Africa","CM":"Africa","ZW":"Africa","ZM":"Africa","MZ":"Africa",
    # Asia
    "CN":"Asia","IN":"Asia","ID":"Asia","JP":"Asia","KR":"Asia","TH":"Asia","VN":"Asia",
    "PH":"Asia","MY":"Asia","SG":"Asia","PK":"Asia","BD":"Asia","LK":"Asia","NP":"Asia",
    "SA":"Asia","AE":"Asia","IL":"Asia","TR":"Asia","IR":"Asia","IQ":"Asia","KZ":"Asia",
    "MM":"Asia","KH":"Asia","LA":"Asia","MN":"Asia","TW":"Asia","HK":"Asia",
    # North America
    "US":"North America","CA":"North America","MX":"North America","GT":"North America",
    "CU":"North America","DO":"North America","HN":"North America","CR":"North America",
    "PA":"North America","NI":"North America","SV":"North America","JM":"North America",
    "HT":"North America","BZ":"North America","BS":"North America",
    # South America
    "BR":"South America","AR":"South America","CO":"South America","CL":"South America",
    "PE":"South America","VE":"South America","EC":"South America","BO":"South America",
    "PY":"South America","UY":"South America","GY":"South America","SR":"South America",
    # Europe
    "GB":"Europe","FR":"Europe","DE":"Europe","IT":"Europe","ES":"Europe","PT":"Europe",
    "NL":"Europe","BE":"Europe","CH":"Europe","AT":"Europe","SE":"Europe","NO":"Europe",
    "DK":"Europe","FI":"Europe","IE":"Europe","PL":"Europe","CZ":"Europe","GR":"Europe",
    "RO":"Europe","HU":"Europe","UA":"Europe","RU":"Europe","BG":"Europe","HR":"Europe",
    "RS":"Europe","SK":"Europe","SI":"Europe","LT":"Europe","LV":"Europe","EE":"Europe",
    "IS":"Europe","LU":"Europe",
    # Oceania
    "AU":"Oceania","NZ":"Oceania","FJ":"Oceania","PG":"Oceania","NC":"Oceania",
    "SB":"Oceania","VU":"Oceania","WS":"Oceania","TO":"Oceania",
}

_model = None
_lock = threading.Lock()


def _load_model():
    global _model
    import formulas
    _model = formulas.ExcelModel().loads(MODEL_XLSX).finish()


def _cv(sol, sheet, a1):
    suf = "!" + a1
    for k in sol:
        ku = k.upper()
        if sheet in ku and ku.endswith(suf):
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

    FW, FUND, REV, COST, PARAM, LOC = (_tag("Framework"), _tag("Funding"),
        _tag("Revenues"), _tag("Costs"), _tag("Parameters"), _tag("Location_Data"))

    green, fish, region = inp["chosen_green"], inp["chosen_fish"], inp["region"]

    ov = {
        f"{FW}!M9":  green,
        f"{FW}!M10": fish,
        f"{FW}!M6":  inp["total_area"],
        f"{FW}!M5":  inp["equity"],
        f"{FW}!M11": region,
        f"{FUND}!D5": inp["cost_of_equity"],
        f"{FUND}!D8": inp["debt_interest_rate"],
    }

    # --- Step 3 price/cost overrides (write into the lookup tables) ---
    if inp.get("crop_price") is not None:
        ov[f"{REV}!C{GREEN_ROW_REV[green]}"] = inp["crop_price"]
    if inp.get("fish_price") is not None:
        ov[f"{REV}!C{FISH_ROW_REV[fish]}"] = inp["fish_price"]
    if inp.get("crop_younglings") is not None:
        ov[f"{COST}!E{GREEN_ROW_COST[green]}"] = inp["crop_younglings"]
    if inp.get("substrate") is not None:
        ov[f"{COST}!F{GREEN_ROW_COST[green]}"] = inp["substrate"]
    if inp.get("fish_younglings") is not None:
        ov[f"{COST}!E{FISH_ROW_COST[fish]}"] = inp["fish_younglings"]
    if inp.get("fish_feed") is not None:
        ov[f"{COST}!F{FISH_ROW_COST[fish]}"] = inp["fish_feed"]
    # energy/water: override the active region's avg cell in Parameters table
    rrow = REGION_ROW_PARAM[region]
    if inp.get("energy_price") is not None:
        ov[f"{PARAM}!C{rrow}"] = inp["energy_price"]   # avg electricity col C
    if inp.get("water_price") is not None:
        ov[f"{PARAM}!F{rrow}"] = inp["water_price"]    # avg water col F

    # --- Step 1 climate overrides (write 12 months into Location_Data) ---
    climate = inp.get("climate")
    if climate:
        for i in range(12):
            row = 7 + i
            ov[f"{LOC}!B{row}"] = climate["high"][i]
            ov[f"{LOC}!C{row}"] = climate["mean"][i]
            ov[f"{LOC}!D{row}"] = climate["low"][i]
            ov[f"{LOC}!E{row}"] = climate["precip"][i]
            ov[f"{LOC}!F{row}"] = climate["et"][i]

    sol = _model.calculate(inputs=ov)

    npv = _num(_cv(sol, "FRAMEWORK", "S20"))
    verdict = _cv(sol, "FRAMEWORK", "T20")
    verdict = verdict.strip() if isinstance(verdict, str) else ("YES" if (npv or 0) > 0 else "NO")
    payback = _num(_cv(sol, "FRAMEWORK", "S22"))
    wacc = _num(_cv(sol, "FUNDING", "L5"))
    total_inv = _num(_cv(sol, "INVESTMENT", "K4"))
    gross_rev = _num(_cv(sol, "FRAMEWORK", "U3"))

    cols = ["S","T","U","V","W","X","Y","Z","AA","AB","AC"]
    fcf = [_num(_cv(sol, "FRAMEWORK", c + "16")) or 0 for c in cols]
    acc = [_num(_cv(sol, "FRAMEWORK", c + "18")) or 0 for c in cols]

    wf_labels = ["Gross Revenue","Sales Taxes","COGS","People Costs","Water Costs",
                 "Energy Costs","Debt Interest","Corporate Tax","Investment",
                 "Loan Repayment","Free Cash Flow"]
    wf_vals = [_num(_cv(sol, "FRAMEWORK", f"AF{3+i}")) or 0 for i in range(11)]
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


app = FastAPI(title="Aquaponics Dashboard", version="3.0.0", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


class Climate(BaseModel):
    high: List[float]
    mean: List[float]
    low: List[float]
    precip: List[float]
    et: List[float]


class CalcRequest(BaseModel):
    chosen_green: str = "Lettuce"
    chosen_fish: str = "Tilapia"
    total_area: float = Field(1000, ge=1)
    equity: float = Field(10000, ge=0)
    cost_of_equity: float = Field(0.171, ge=0, le=2)
    debt_interest_rate: float = Field(0.1186, ge=0, le=2)
    region: str = "South America"
    # Step 3 (all optional; fall back to sheet defaults if omitted)
    crop_price: Optional[float] = None
    fish_price: Optional[float] = None
    crop_younglings: Optional[float] = None
    substrate: Optional[float] = None
    fish_younglings: Optional[float] = None
    fish_feed: Optional[float] = None
    energy_price: Optional[float] = None
    water_price: Optional[float] = None
    # Step 1
    climate: Optional[Climate] = None


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/options")
def options():
    return dict(
        greens=GREENS, fish=FISH, regions=REGIONS,
        green_price_default=GREEN_PRICE_DEFAULT, fish_price_default=FISH_PRICE_DEFAULT,
        green_cost_default=GREEN_COST_DEFAULT, fish_cost_default=FISH_COST_DEFAULT,
        region_defaults=REGION_DEFAULTS, country_region=COUNTRY_REGION,
        defaults=dict(chosen_green="Lettuce", chosen_fish="Tilapia", total_area=1000,
                      equity=10000, cost_of_equity=0.171, debt_interest_rate=0.1186,
                      region="South America"),
    )


@app.get("/region_for_country/{code}")
def region_for_country(code: str):
    return {"region": COUNTRY_REGION.get(code.upper(), "South America")}


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
