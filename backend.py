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
from typing import Optional, List, Dict

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

# ── Equipment with a productivity factor (Investment sheet) ──────────────────
# Toggle writes 1/0 into Investment!C{row}. Column G is that item's productivity
# factor; the system factor is MAX(activated) = Investment!H7, which multiplies
# crop revenue. Items with factor 1.0 are cost/energy-only (no yield boost).
EQUIPMENT = [
    dict(row=25, name="Aeration Pump",              group="Fish tank",       factor=1.05, default=0),
    dict(row=35, name="Ventilators",                group="Hydroponic area", factor=1.00, default=1),
    dict(row=36, name="De-humidifier",              group="Hydroponic area", factor=1.00, default=1),
    dict(row=37, name="Greenhouse",                 group="Hydroponic area", factor=1.05, default=0),
    dict(row=38, name="Lighting",                   group="Hydroponic area", factor=0.98, default=0),
    dict(row=43, name="pH sensor",                  group="Hydroponic unit", factor=1.02, default=1),
    dict(row=44, name="Nutrient dosing unit",       group="Hydroponic unit", factor=1.00, default=0),
    dict(row=45, name="Hanging/moving Gutter",      group="Hydroponic unit", factor=1.00, default=0),
    dict(row=46, name="UV-sterilization",           group="Hydroponic unit", factor=1.00, default=0),
    dict(row=47, name="Pollination vibrator",       group="Hydroponic unit", factor=1.00, default=0),
    dict(row=51, name="Biologic Surface Area",      group="Synergies",       factor=1.00, default=1),
    dict(row=52, name="Periphyton Surface Area",    group="Synergies",       factor=1.01, default=1),
    dict(row=53, name="Biofloc",                    group="Synergies",       factor=1.02, default=0),
    dict(row=54, name="Aerobic mineralization",     group="Synergies",       factor=1.11, default=1),
    dict(row=55, name="UASB Bioreactor",            group="Synergies",       factor=1.00, default=0),
    dict(row=56, name="EGSB Bioreactor",            group="Synergies",       factor=1.00, default=0),
    dict(row=57, name="Gravity Separation (RFS)",   group="Synergies",       factor=1.00, default=0),
]
EQUIP_ROWS = {e["row"] for e in EQUIPMENT}

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

# ── "Advanced" red-font user inputs, grouped for the modal ───────────────────
# Each entry: key -> (sheet, cell-template, label, group, step, unit)
# {G} = row of the chosen crop on System_Parameters, {F} = row of chosen fish.
SP_GREEN_ROW = {g: 19 + i for i, g in enumerate(GREENS)}
SP_FISH_ROW = {"Tilapia": 33, "Trout": 34}

ADVANCED = [
    # System sizing
    dict(key="pfrm",            sheet="System_Parameters", cell="E4",  label="Plant : fish mass ratio", group="System sizing", step=0.1,   unit=""),
    dict(key="tank_volume",     sheet="System_Parameters", cell="E5",  label="Avg tank volume",         group="System sizing", step=1,     unit="m³"),
    dict(key="tank_height",     sheet="System_Parameters", cell="E6",  label="Tank height",             group="System sizing", step=0.1,   unit="m"),
    dict(key="max_hydro_unit",  sheet="System_Parameters", cell="E7",  label="Max hydroponic unit",     group="System sizing", step=1,     unit="m²"),
    dict(key="admin_pct",       sheet="System_Parameters", cell="E8",  label="Admin / service area",    group="System sizing", step=0.01,  unit="fraction"),
    # Crop biology (row depends on chosen crop)
    dict(key="crop_lifecycle",  sheet="System_Parameters", cell="C{G}", label="Lifecycle",              group="Crop biology",  step=1,     unit="days"),
    dict(key="crop_density",    sheet="System_Parameters", cell="D{G}", label="Plants per m²",          group="Crop biology",  step=1,     unit="plants/m²"),
    dict(key="crop_kg",         sheet="System_Parameters", cell="E{G}", label="Yield per plant",        group="Crop biology",  step=0.01,  unit="kg"),
    dict(key="crop_tmin",       sheet="System_Parameters", cell="F{G}", label="Min temperature",        group="Crop biology",  step=1,     unit="°C"),
    dict(key="crop_tmax",       sheet="System_Parameters", cell="G{G}", label="Max temperature",        group="Crop biology",  step=1,     unit="°C"),
    dict(key="crop_tideal",     sheet="System_Parameters", cell="H{G}", label="Ideal temperature",      group="Crop biology",  step=1,     unit="°C"),
    # Fish biology
    dict(key="fish_lifecycle",  sheet="System_Parameters", cell="C{F}", label="Lifecycle",              group="Fish biology",  step=1,     unit="days"),
    dict(key="fish_density",    sheet="System_Parameters", cell="D{F}", label="Fish per m³",            group="Fish biology",  step=1,     unit="fish/m³"),
    dict(key="fish_kg",         sheet="System_Parameters", cell="E{F}", label="Weight at harvest",      group="Fish biology",  step=0.05,  unit="kg"),
    dict(key="fish_tmin",       sheet="System_Parameters", cell="F{F}", label="Min temperature",        group="Fish biology",  step=1,     unit="°C"),
    dict(key="fish_tmax",       sheet="System_Parameters", cell="G{F}", label="Max temperature",        group="Fish biology",  step=1,     unit="°C"),
    dict(key="fish_tideal",     sheet="System_Parameters", cell="H{F}", label="Ideal temperature",      group="Fish biology",  step=1,     unit="°C"),
    # Revenue assumptions
    dict(key="crop_loss",       sheet="Revenues", cell="B8",  label="Crop loss rate",        group="Revenue",  step=0.01, unit="fraction"),
    dict(key="fish_loss",       sheet="Revenues", cell="B9",  label="Fish loss rate",        group="Revenue",  step=0.01, unit="fraction"),
    dict(key="first_year_ramp", sheet="Revenues", cell="B10", label="Year-1 ramp",           group="Revenue",  step=0.05, unit="fraction"),
    dict(key="sales_tax",       sheet="Revenues", cell="B11", label="Sales tax",             group="Revenue",  step=0.005,unit="fraction"),
    dict(key="corp_tax",        sheet="Revenues", cell="B12", label="Corporate tax",         group="Revenue",  step=0.01, unit="fraction"),
    dict(key="services_a",      sheet="Revenues", cell="B13", label="Services revenue A",    group="Revenue",  step=100,  unit="$/yr"),
    dict(key="services_b",      sheet="Revenues", cell="B14", label="Services revenue B",    group="Revenue",  step=100,  unit="$/yr"),
    # Investment & financing
    dict(key="maintenance_pct", sheet="Investment", cell="B6", label="Maintenance",          group="Investment & financing", step=0.005, unit="of capex"),
    dict(key="spare_parts_pct", sheet="Investment", cell="B7", label="Spare parts",          group="Investment & financing", step=0.005, unit="of capex"),
    dict(key="grace_years",     sheet="Funding",    cell="E8", label="Debt grace period",    group="Investment & financing", step=1,     unit="years"),
    dict(key="debt_years",      sheet="Funding",    cell="F8", label="Debt term",            group="Investment & financing", step=1,     unit="years"),
    # Operating costs
    dict(key="crop_worker",     sheet="Costs", cell="B21", label="Crop worker salary",       group="Operating costs", step=50, unit="$/month"),
    dict(key="aqua_worker",     sheet="Costs", cell="B22", label="Aquaculture worker salary",group="Operating costs", step=50, unit="$/month"),
    dict(key="biologic_surface",sheet="Costs", cell="B26", label="Biologic surface area",    group="Operating costs", step=5,  unit="$/tank"),
    dict(key="nutrients",       sheet="Costs", cell="B32", label="Nutrients",                group="Operating costs", step=0.01,unit="$/plant"),
    dict(key="controls",        sheet="Costs", cell="B34", label="Controls",                 group="Operating costs", step=1,  unit="$/m²"),
]

# Defaults so the modal can pre-fill without a round trip.
GREEN_BIO = {
    "Tomato":                 dict(crop_lifecycle=365, crop_density=3.9, crop_kg=10,      crop_tmin=13, crop_tmax=32, crop_tideal=23),
    "Lettuce":                dict(crop_lifecycle=35,  crop_density=25,  crop_kg=0.2,     crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Chicória":               dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.13232, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Almeirão Pão de Açúcar": dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.09816, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Almeirão Amargo":        dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.11576, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Espinafre":              dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.07872, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Alface":                 dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.07872, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Salsinha":               dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.03587, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Manjericão":             dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.04206, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Agrião":                 dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.07369, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
    "Cebolinha":              dict(crop_lifecycle=28,  crop_density=16,  crop_kg=0.02525, crop_tmin=7,  crop_tmax=24, crop_tideal=18),
}
FISH_BIO = {
    "Tilapia": dict(fish_lifecycle=180, fish_density=100, fish_kg=0.6, fish_tmin=11, fish_tmax=35, fish_tideal=29),
    "Trout":   dict(fish_lifecycle=365, fish_density=22,  fish_kg=0.5, fish_tmin=0,  fish_tmax=21, fish_tideal=14),
}
STATIC_ADV_DEFAULTS = dict(
    pfrm=3.1, tank_volume=10, tank_height=1.2, max_hydro_unit=5, admin_pct=0.15,
    crop_loss=0.05, fish_loss=0.05, first_year_ramp=0.6, sales_tax=0.065, corp_tax=0.15,
    services_a=0, services_b=0, maintenance_pct=0.01, spare_parts_pct=0.025,
    grace_years=0, debt_years=10, crop_worker=1500, aqua_worker=3000,
    biologic_surface=50, nutrients=0.15, controls=10,
)

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

    # --- Equipment toggles: write 1/0 into Investment!C{row} ---
    equip = inp.get("equipment") or {}
    if equip:
        INV = _tag("Investment")
        for row_str, on in equip.items():
            try:
                row = int(row_str)
            except (TypeError, ValueError):
                continue
            if row in EQUIP_ROWS:
                ov[f"{INV}!C{row}"] = 1 if on else 0

    # --- Advanced (red-font) overrides ---
    advanced = inp.get("advanced") or {}
    if advanced:
        grow = SP_GREEN_ROW.get(green)
        frow = SP_FISH_ROW.get(fish)
        spec = {a["key"]: a for a in ADVANCED}
        for key, val in advanced.items():
            a = spec.get(key)
            if a is None or val is None:
                continue
            cell = a["cell"].replace("{G}", str(grow)).replace("{F}", str(frow))
            ov[f"{_tag(a['sheet'])}!{cell}"] = val

    sol = _model.calculate(inputs=ov)

    npv = _num(_cv(sol, "FRAMEWORK", "S20"))
    verdict = _cv(sol, "FRAMEWORK", "T20")
    verdict = verdict.strip() if isinstance(verdict, str) else ("YES" if (npv or 0) > 0 else "NO")
    payback = _num(_cv(sol, "FRAMEWORK", "S22"))
    wacc = _num(_cv(sol, "FUNDING", "L5"))
    total_inv = _num(_cv(sol, "INVESTMENT", "K4"))
    gross_rev = _num(_cv(sol, "FRAMEWORK", "U3"))
    prod_factor = _num(_cv(sol, "INVESTMENT", "H7"))
    rev_greens   = _num(_cv(sol, "REVENUES", "G4")) or 0.0
    rev_fish     = _num(_cv(sol, "REVENUES", "G6")) or 0.0
    rev_services = _num(_cv(sol, "REVENUES", "G7")) or 0.0

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
        productivity_factor=round(prod_factor, 4) if prod_factor is not None else None,
        revenue_split=dict(greens=round(rev_greens, 2), fish=round(rev_fish, 2),
                           services=round(rev_services, 2)),
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
    # Equipment toggles: {"35": 1, "37": 0, ...} keyed by Investment sheet row
    equipment: Optional[Dict[str, int]] = None
    # Advanced red-font inputs: {"crop_density": 20, "sales_tax": 0.07, ...}
    advanced: Optional[Dict[str, float]] = None
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
        equipment=EQUIPMENT,
        advanced=ADVANCED, advanced_defaults=STATIC_ADV_DEFAULTS,
        green_bio=GREEN_BIO, fish_bio=FISH_BIO,
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
