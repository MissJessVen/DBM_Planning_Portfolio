#!/usr/bin/env python3
"""
OTB & Replenishment Planning Portfolio - end-to-end pipeline
=============================================================

Rebuilds every number, chart and file behind the portfolio dashboard from the
raw PhilGEPS DepEd zip (19 quarterly CSVs, 2021-2025):

  1. Load, normalize and de-duplicate the quarterly files (data-quality log)
  2. Map free-text goods lines to 39 SKUs / 9 categories / 2 channels
       Quick-Buy (DTC stand-in)  = small-value + shopping procurement
       Bid       (Amazon stand-in) = public bidding and other formal modes
  3. Carve tender-size orders (> channel P99) out of base demand
  4. Build the monthly SKU x channel demand grid (Apr-2022 .. Dec-2025)
  5. Forecast models v1.0 / v1.1 / v1.2a / v1.2b / v1.2c + rolling back-test
  6. OTB engine: safety stock, EOM targets, end-of-life rule, Jul-Dec 2025
     inventory simulation, Jan-Mar 2026 plan
  7. Top-down reconciliation and analyst override (scale cells > 1.5x TD to 1.25x TD)
  8. SKU line review flags (stockout, overstock, EOL, lumpy, >2.5x LY, moved >30%)
  9. Launch-and-fade curve and SKU lifecycle table
 10. Charts (PNG), JSON/CSV datasets and the Excel OTB planner

Usage
-----
    python otb_portfolio.py --zip COMBINED_PHILGEPS_DEPED.zip --out output

Requirements: pandas, numpy, matplotlib, openpyxl   (see requirements.txt)
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm


# =============================================================================
# 0. CONFIGURATION - every planning lever lives here
# =============================================================================
@dataclass(frozen=True)
class Config:
    model_start: str = "2022-04-01"          # 2022-Q1 file is missing -> start after the gap
    model_end: str = "2025-12-01"
    train_end_for_cap: str = "2025-06-01"    # tender P99 learned on data up to here
    tender_pct: float = 0.99                 # lines above channel P99 = tender pipeline
    si_shrink: float = 0.6                   # seasonal index = 1 + 0.6*(raw-1)  (40% shrink to flat)
    backtest_first_origin: str = "2024-07-01"
    backtest_last_origin: str = "2025-10-01"
    horizons: tuple = (1, 2, 3)
    z: float = 1.28                          # 90% cycle service level
    lead_days: dict = None                   # filled from data medians below
    ss_cap_months: dict = None               # SS capped at N months of next-month forecast
    fade_min_prior6: float = 2000.0          # PHP k; EOL rule only on material series
    fade_ratio_vs_channel: float = 0.5       # EOL if SKU 6m ratio < 0.5 x channel 6m ratio
    sim_start: str = "2025-07-01"            # inventory simulation start
    plan_months: tuple = ("2026-01", "2026-02", "2026-03")
    override_trigger: float = 1.5            # BU > 1.5 x TD -> override
    override_target: float = 1.25            # scale to 1.25 x TD
    cost_ratio: float = 0.70                 # buy cost / line value (assumption)
    gate_wape_pts: float = 2.0               # ship gate: WAPE must improve >= 2 pts
    gate_bias_abs: float = 10.0              # ... and |bias| <= 10%


CFG = Config(lead_days={"Quick-Buy": 14, "Bid": 51},
             ss_cap_months={"Quick-Buy": 1.0, "Bid": 0.5})

# Chart palette (validated reference categorical slots + recessive ink)
C_BLUE, C_ORANGE, C_AQUA = "#257de0", "#f47230", "#00b894"
C_MUTED, C_GRID, C_TEXT, C_TEXT2 = "#bdbdbd", "#e6e6e6", "#191919", "#666666"
C_BAD, C_OK = "#dc3c3d", "#0aac23"


def setup_matplotlib():
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Arial", "Helvetica", "DejaVu Sans"],
        "font.size": 10, "axes.titlesize": 12, "axes.titleweight": "bold",
        "axes.edgecolor": C_GRID, "axes.labelcolor": C_TEXT2,
        "xtick.color": C_TEXT2, "ytick.color": C_TEXT2,
        "axes.grid": True, "grid.color": C_GRID, "grid.linewidth": 0.8,
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.axisbelow": True, "figure.dpi": 110, "savefig.dpi": 160,
        "savefig.bbox": "tight", "legend.frameon": False,
    })


# =============================================================================
# 1. LOAD, NORMALIZE, DE-DUPLICATE
# =============================================================================
NUM_COLS = ["approved_budget_php_thousands", "quantity", "item_budget_php_thousands",
            "contract_amount_php_thousands", "line_item_no", "award_reference_no",
            "bid_reference_no", "contract_duration", "unspsc_code"]
DEDUP_KEY = ["procuring_entity", "procurement_reference_date", "notice_title", "line_item_no",
             "item_name", "quantity", "item_budget_php_thousands", "award_reference_no",
             "contract_amount_php_thousands", "awardee_organization_name"]


def extract_zip(zip_path: str) -> str:
    """Extract into a fresh temp dir (untrusted archive stays isolated)."""
    out = tempfile.mkdtemp(prefix="philgeps_")
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(out)
    return out


def load_raw(folder: str) -> pd.DataFrame:
    files = sorted(glob.glob(os.path.join(folder, "**", "*.csv"), recursive=True))
    if not files:
        raise FileNotFoundError(f"No CSV files under {folder}")
    frames = []
    for f in files:
        d = pd.read_csv(f, low_memory=False, dtype=str)
        d["src_file"] = os.path.basename(f)
        frames.append(d)
    return pd.concat(frames, ignore_index=True)


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    dq = {"raw_rows": len(df), "files": df.src_file.nunique()}
    # same value is written as 3840 / 3840.0 / NULL / '' across files -> normalize first
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c].replace({"NULL": np.nan, "": np.nan}), errors="coerce")
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].replace({"NULL": np.nan}).str.strip()
    df["is_deped"] = df.is_deped.str.upper()                  # True vs TRUE
    df["date"] = pd.to_datetime(df.procurement_reference_date, errors="coerce")
    fq = df.src_file.str.extract(r"(\d{4})_?\s?Q(\d)")
    df["file_q"] = fq[0] + "-Q" + fq[1]
    dq["rows_outside_file_quarter"] = int((df.file_q != df.fiscal_quarter).sum())
    before = len(df)
    d = df.drop_duplicates(subset=DEDUP_KEY).copy()
    dq["exact_duplicates_removed"] = before - len(d)
    dq["clean_rows"] = len(d)
    dq["rows_by_quarter"] = d.fiscal_quarter.value_counts().sort_index().to_dict()
    dq["missing_item_budget_share"] = round(d.item_budget_php_thousands.isna().mean(), 3)
    return d, dq


# =============================================================================
# 2. SKU MASTER: keyword rules -> mixed-lot fallbacks by business category
# =============================================================================
SERVICE_RX = (r"catering|meal|snack|lodging|accommodation|venue|hotel|food and|board and|lease|"
              r"rental|transport|repair|construction|printing services|reproduction")
SERVICE_BC = ("Catering|Hotel|Lodging|Lease|Travel|Transportation|Construction Projects|"
              "Printing Services|Repair")
SKUS = [
    ("Office & Print", "OP-01", "Bond paper (A4/long/short)", r"bond paper|copy paper|\bream"),
    ("Office & Print", "OP-02", "Ink & ink refills", r"\bink\b|ink refill|ink bottle|ink cartridge"),
    ("Office & Print", "OP-03", "Toner cartridges", r"toner"),
    ("Office & Print", "OP-04", "Folders, envelopes & filing", r"folder|envelope|fastener|filing"),
    ("Office & Print", "OP-05", "Writing instruments", r"ballpen|ball pen|pencil|marker|highlighter"),
    ("Office & Print", "OP-06", "Tarpaulin & signage", r"tarpaulin|tarp\b|signage|streamer"),
    ("IT & Devices", "IT-01", "Laptops", r"laptop|notebook computer"),
    ("IT & Devices", "IT-02", "Desktop computers", r"desktop|computer set|personal computer|all-in-one"),
    ("IT & Devices", "IT-03", "Printers & scanners", r"printer(?! ink)|scanner"),
    ("IT & Devices", "IT-04", "Tablets", r"tablet(?!s? \d+ ?mg)"),
    ("IT & Devices", "IT-05", "Projectors, TVs & AV", r"projector|smart tv|television|\btv\b|speaker|microphone|sound system"),
    ("IT & Devices", "IT-06", "Storage & peripherals", r"flash drive|usb|external hard|hard drive|ssd|mouse|keyboard|router|ups\b|avr\b"),
    ("Food & Nutrition", "FN-01", "Rice", r"\brice\b"),
    ("Food & Nutrition", "FN-02", "Milk (fresh/UHT/powder)", r"milk"),
    ("Food & Nutrition", "FN-03", "Bread & baked goods", r"bread|pandesal|bakery|cookies|biscuit"),
    ("Food & Nutrition", "FN-04", "Fruits, vegetables & eggs", r"fruit|vegetable|banana|egg\b|eggs"),
    ("Food & Nutrition", "FN-05", "Canned & processed food", r"canned|noodle|sardine|corned|food pack"),
    ("Learning Materials", "LM-01", "Learning modules / SLMs", r"module|self-learning|\bslm"),
    ("Learning Materials", "LM-02", "Books & supplementary readers", r"\bbook|reader|textbook|dictionar"),
    ("Learning Materials", "LM-03", "Science & math kits", r"science kit|math kit|laboratory|microscope|manipulative"),
    ("Learning Materials", "LM-04", "Art & classroom supplies", r"crayon|art paper|cartolina|manila paper|construction paper|art material"),
    ("Health & Hygiene", "HH-01", "Alcohol & sanitizers", r"alcohol|sanitizer"),
    ("Health & Hygiene", "HH-02", "Face masks & PPE", r"face mask|facemask|surgical mask|kn95|ppe|face shield"),
    ("Health & Hygiene", "HH-03", "Vitamins & medicines", r"vitamin|paracetamol|medicine|drug|amoxicillin|deworm"),
    ("Health & Hygiene", "HH-04", "Cleaning & janitorial", r"detergent|disinfect|bleach|soap|tissue|trash bag|garbage bag|broom|mop|janitorial|cleaning"),
    ("Health & Hygiene", "HH-05", "Medical & first-aid supplies", r"first aid|bandage|gloves|thermometer|blood pressure|medical suppl"),
    ("Events & Awards", "EA-01", "Medals, trophies & plaques", r"medal|trophy|trophies|plaque|certificate holder|ribbon"),
    ("Events & Awards", "EA-02", "Event kits & giveaways", r"kit\b|kits\b|giveaway|bag\b|bags\b|lanyard|id lace"),
    ("Apparel & Sports", "AS-01", "Uniforms & t-shirts", r"uniform|t-shirt|tshirt|shirt|jersey|polo"),
    ("Apparel & Sports", "AS-02", "Sports equipment", r"\bball\b|basketball|volleyball|sports equipment|badminton|chess|table tennis|sepak|athletic"),
    ("Furniture & Fixtures", "FF-01", "Chairs & armchairs", r"chair"),
    ("Furniture & Fixtures", "FF-02", "Tables, desks & cabinets", r"table(?! tennis)|desk|cabinet|shelf|shelves|rack"),
    ("Furniture & Fixtures", "FF-03", "Fans, aircon & electricals", r"electric fan|\bfan\b|aircon|air con|air-con|bulb|led light|extension|wire\b|electrical"),
]
EXTRA_RULES = [
    ("OP-05", "Writing instruments", "Office & Print", r"sign pen|correction tape|paper clip|\btape\b|stapler|staple wire"),
    ("FN-03", "Bread & baked goods", "Food & Nutrition", r"nutribun"),
]
MIXED_LOT_BY_BC = {
    "Office Equipment Supplies and Consumables": ("OP-99", "Office supplies – mixed lots", "Office & Print"),
    "Office Supplies and Devices": ("OP-99", "Office supplies – mixed lots", "Office & Print"),
    "Printing Supplies": ("OP-99", "Office supplies – mixed lots", "Office & Print"),
    "Office Equipment": ("IT-99", "Office/IT equipment – mixed lots", "IT & Devices"),
    "Information Technology": ("IT-99", "Office/IT equipment – mixed lots", "IT & Devices"),
    "Information Technology Parts & Accessories & Perip": ("IT-99", "Office/IT equipment – mixed lots", "IT & Devices"),
    "Food Stuff": ("FN-99", "Food items – mixed lots", "Food & Nutrition"),
    "Dairy Products": ("FN-02", "Milk (fresh/UHT/powder)", "Food & Nutrition"),
    "Educational Materials and Supplies": ("LM-99", "Learning materials – mixed lots", "Learning Materials"),
    "Books, Maps and Other Publications": ("LM-02", "Books & supplementary readers", "Learning Materials"),
    "Medical Supplies and Laboratory Instrument": ("HH-99", "Health & hygiene – mixed lots", "Health & Hygiene"),
    "Janitorial Supplies": ("HH-04", "Cleaning & janitorial", "Health & Hygiene"),
    "Sporting Goods": ("AS-02", "Sports equipment", "Apparel & Sports"),
    "Garments": ("AS-01", "Uniforms & t-shirts", "Apparel & Sports"),
    "Tokens and Awards": ("EA-01", "Medals, trophies & plaques", "Events & Awards"),
    "Furniture": ("FF-02", "Tables, desks & cabinets", "Furniture & Fixtures"),
    "General Merchandise": ("GM-99", "General merchandise – mixed lots", "General Merchandise"),
}


def map_skus(d: pd.DataFrame) -> pd.DataFrame:
    d = d[d.classification.str.startswith("Goods", na=False)].copy()
    # line value: item budget, else notice budget split evenly across its lines
    n_lines = d.groupby(["procuring_entity", "procurement_reference_date", "notice_title"])["item_name"].transform("size")
    d["value_k"] = d.item_budget_php_thousands.fillna(d.approved_budget_php_thousands / n_lines)
    it = (d.item_name.fillna("") + " | " + d.item_description.fillna("")).str.lower()
    d["svc"] = it.str.contains(SERVICE_RX) | d.business_category.str.contains(SERVICE_BC, na=False)
    d["sku"] = None; d["sku_name"] = None; d["category"] = None
    for cat, code, name, rx in SKUS:
        m = d.sku.isna() & (~d.svc) & it.str.contains(rx, regex=True)
        d.loc[m, ["sku", "sku_name", "category"]] = [code, name, cat]
    for code, name, cat, rx in EXTRA_RULES:
        m = d.sku.isna() & (~d.svc) & it.str.contains(rx)
        d.loc[m, ["sku", "sku_name", "category"]] = [code, name, cat]
    for bc, (code, name, cat) in MIXED_LOT_BY_BC.items():
        m = d.sku.isna() & (~d.svc) & (d.business_category == bc)
        d.loc[m, ["sku", "sku_name", "category"]] = [code, name, cat]
    d["chan"] = np.where(d.procurement_mode.str.contains("Small Value|Shopping", na=False), "Quick-Buy", "Bid")
    a = d[d.sku.notna() & d.value_k.notna() & (d.value_k > 0)].copy()
    a["month"] = a.date.dt.to_period("M").dt.to_timestamp()
    return a


# =============================================================================
# 3-4. TENDER CARVE-OUT + MONTHLY DEMAND GRID
# =============================================================================
class DemandModel:
    """Holds the SKU x channel monthly grid and every forecasting / OTB routine."""

    def __init__(self, lines: pd.DataFrame, cfg: Config = CFG):
        self.cfg = cfg
        a = lines[lines.month >= cfg.model_start].copy()
        tr = a[a.month <= cfg.train_end_for_cap]
        self.cap = tr.groupby("chan").value_k.quantile(cfg.tender_pct).to_dict()
        a["tender"] = a.value_k > a.chan.map(self.cap)
        self.L = a
        self.months = pd.date_range(cfg.model_start, cfg.model_end, freq="MS")
        keys = a[["sku", "sku_name", "category"]].drop_duplicates("sku")
        base = a[~a.tender]
        g = (base.groupby(["sku", "chan", "month"]).value_k.sum().unstack("month")
             .reindex(columns=self.months, fill_value=0).fillna(0))
        idx = pd.MultiIndex.from_product([keys.sku.unique(), ["Quick-Buy", "Bid"]], names=["sku", "chan"])
        self.grid = g.reindex(idx, fill_value=0)
        self.cat_of = keys.set_index("sku").category
        self.name_of = keys.set_index("sku").sku_name
        self.Y = self.grid.values
        self.M = len(self.months)
        self.chans = np.array([c for _, c in self.grid.index])
        self.cats = np.array([self.cat_of[k] for k, _ in self.grid.index])
        self.lt = np.array([cfg.lead_days[c] / 30 for c in self.chans])
        self.ss_cap = np.array([cfg.ss_cap_months[c] for c in self.chans])
        self._si_cache = {}
        self.sig = np.zeros(len(self.grid))

    # ---- seasonality: category x month-of-year index, shrunk toward 1 -------
    def si_table(self, upto: int) -> pd.DataFrame:
        if upto in self._si_cache:
            return self._si_cache[upto]
        g = self.grid.iloc[:, :upto].groupby(level="sku").sum()
        g = g.groupby(g.index.map(self.cat_of)).sum()
        tot = g.T
        tot["moy"] = tot.index.month
        out = {}
        for c in g.index:
            s = tot[[c, "moy"]].copy()
            s["r"] = s[c] / s[c].rolling(12, min_periods=6, center=True).mean()
            r = s.groupby("moy").r.mean().reindex(range(1, 13)).fillna(1)
            r = r / r.mean()
            out[c] = 1 + self.cfg.si_shrink * (r - 1)
        self._si_cache[upto] = pd.DataFrame(out)
        return self._si_cache[upto]

    def _season_vec(self, t: int, h: int) -> np.ndarray:
        si = self.si_table(t)
        tgt = (self.months[0] + pd.DateOffset(months=t + h - 1)).month
        return np.array([si.loc[tgt, c] for c in self.cats])

    # ---- 5. forecast versions ------------------------------------------------
    def forecast(self, t: int, h: int, ver: str) -> np.ndarray:
        """Forecast month t+h-1 using only data before index t."""
        hist = self.Y[:, :t]
        if ver == "v1.0":                                 # trailing 3-month average
            return hist[:, -3:].mean(1)
        s = self._season_vec(t, h)
        if ver == "v1.2c":                                # outlier months capped at 3x median
            hh = hist[:, -12:].copy()
            med = np.median(hh, axis=1, keepdims=True); mean = hh.mean(1, keepdims=True)
            capv = np.maximum(3.0 * np.maximum(med, 0.5 * mean), 1e-9)
            return np.minimum(hh, capv).mean(1) * s
        f = hist[:, -12:].mean(1) * s                     # v1.1: 12-mo level x seasonality
        if ver == "v1.1":
            return f
        if ver == "v1.2a":                                # SKU trend: last 6m vs same 6m LY
            r6 = hist[:, -6:].sum(1); p6 = hist[:, -18:-12].sum(1)
            tr = np.where(p6 > 0, r6 / np.maximum(p6, 1e-9), 1.0)
            tr = np.clip(tr, 0.4, 1.6) ** 0.5
            tr = np.where((r6 + p6) < 50, 1.0, tr)
            return f * tr
        if ver == "v1.2b":                                # channel-level trend, half strength
            out = f.copy()
            for c in np.unique(self.chans):
                m = self.chans == c
                r6 = hist[m, -6:].sum(); p6 = hist[m, -18:-12].sum()
                out[m] *= np.clip(r6 / p6 if p6 > 0 else 1, 0.5, 1.5) ** 0.5
            return out
        raise ValueError(ver)

    def backtest(self, versions=("v1.0", "v1.1", "v1.2a", "v1.2b", "v1.2c")) -> pd.DataFrame:
        rows = []
        origins = [self.months.get_loc(pd.Timestamp(x)) for x in
                   pd.date_range(self.cfg.backtest_first_origin, self.cfg.backtest_last_origin, freq="MS")]
        for t in origins:
            for h in self.cfg.horizons:
                if t + h - 1 >= self.M:
                    continue
                act = self.Y[:, t + h - 1]
                for v in versions:
                    f = self.forecast(t, h, v)
                    for i, (k, c) in enumerate(self.grid.index):
                        rows.append((v, self.months[t].strftime("%Y-%m"), h,
                                     self.months[t + h - 1].strftime("%Y-%m"), k, c, f[i], act[i]))
        bt = pd.DataFrame(rows, columns=["ver", "origin", "h", "target", "sku", "chan", "fc", "act"])
        # forecast-error sigma per series (v1.1, 1-month ahead) feeds safety stock
        e = bt[(bt.ver == "v1.1") & (bt.h == 1)].assign(err=lambda x: x.act - x.fc)
        self.sig = e.groupby(["sku", "chan"]).err.std().reindex(self.grid.index).fillna(0).values
        return bt

    # ---- 6. OTB engine ----------------------------------------------------------
    def fc11(self, t: int, h: int) -> np.ndarray:
        return self.forecast(t, h, "v1.1")

    def fading(self, t: int) -> np.ndarray:
        """End-of-life: last 6m < 50% of same 6m a year earlier, relative to channel trend."""
        hist = self.Y[:, :t]
        r6 = hist[:, -6:].sum(1); p6 = hist[:, -18:-12].sum(1)
        mk = {c: hist[self.chans == c, -6:].sum() / hist[self.chans == c, -18:-12].sum()
              for c in ("Bid", "Quick-Buy")}
        m = np.array([mk[c] for c in self.chans])
        return (p6 > self.cfg.fade_min_prior6) & (r6 < self.cfg.fade_ratio_vs_channel * m * p6)

    def target_eom(self, t: int, h: int, fade: np.ndarray) -> np.ndarray:
        nxt = self.fc11(t, h + 1)
        ss = np.minimum(self.cfg.z * self.sig * np.sqrt(self.lt), self.ss_cap * nxt)
        return np.where(fade, 0, nxt * self.lt + ss)

    def simulate_inventory(self) -> tuple[list, np.ndarray]:
        """Run the planner's own buys Jul-Dec 2025 against actual demand -> Jan-2026 opening stock."""
        t0 = self.months.get_loc(pd.Timestamp(self.cfg.sim_start))
        inv = self.target_eom(t0 - 1, 1, self.fading(t0 - 1))
        sim = []
        for t in range(t0, self.M):
            fd = self.fading(t); ps = self.fc11(t, 1); eom = self.target_eom(t, 1, fd)
            otb = np.maximum(ps + eom - inv, 0)
            end = np.maximum(inv + otb - self.Y[:, t], 0)
            lost = np.maximum(self.Y[:, t] - (inv + otb), 0)
            sim.append(dict(month=self.months[t].strftime("%Y-%m"), bom=inv.sum(), otb=otb.sum(),
                            sales_plan=ps.sum(), act=self.Y[:, t].sum(), eom=end.sum(), lost=lost.sum()))
            inv = end
        return sim, inv

    def plan(self, bom: np.ndarray) -> pd.DataFrame:
        """Jan-Mar 2026 model plan from an origin after the last actual month."""
        rows = []; fd = self.fading(self.M); inv = bom.copy()
        for h, month in enumerate(self.cfg.plan_months, 1):
            ps = self.fc11(self.M, h); eom = self.target_eom(self.M, h, fd)
            otb = np.maximum(ps + eom - inv, 0)
            for i, (k, c) in enumerate(self.grid.index):
                rows.append(dict(month=month, sku=k, sku_name=self.name_of[k], category=self.cat_of[k], chan=c,
                                 bom=inv[i], sales_plan=ps[i], eom_target=eom[i], otb=otb[i], fading=bool(fd[i])))
            inv = inv + otb - ps
        return pd.DataFrame(rows)

    # ---- 7. top-down plan + analyst override ------------------------------------
    def topdown(self, P: pd.DataFrame) -> dict:
        G = self.grid
        cc = pd.DataFrame({"chan": self.chans, "cat": self.cats,
                           "h25": G.loc[:, "2025-07-01":"2025-12-01"].sum(axis=1).values,
                           "h24": G.loc[:, "2024-07-01":"2024-12-01"].sum(axis=1).values})
        mom = cc.groupby(["chan", "cat"])[["h25", "h24"]].sum()
        mom["m"] = np.clip(mom.h25 / mom.h24, 0.5, 1.5) ** 0.5
        TD = {}
        for month, g in P.groupby("month"):
            mo = pd.Timestamp(month).month
            ly_cols = [pd.Timestamp(y, mo, 1) for y in (2023, 2024, 2025)]
            for (c, cat), x in g.groupby(["chan", "category"]):
                ly = G.loc[[(k, c) for k in x.sku], ly_cols].sum().mean()
                TD[(month, c, cat)] = ly * mom.loc[(c, cat), "m"]
        return TD

    def apply_override(self, P: pd.DataFrame, TD: dict) -> pd.DataFrame:
        cfg = self.cfg
        bu = P.groupby(["month", "chan", "category"]).sales_plan.transform("sum")
        P = P.copy()
        P["scale"] = [min(1.0, cfg.override_target * TD[(m, c, cat)] / max(b, 1e-9))
                      if b > cfg.override_trigger * TD[(m, c, cat)] else 1.0
                      for m, c, cat, b in zip(P.month, P.chan, P.category, bu)]
        P["rec_sales"] = P.sales_plan * P.scale
        nxt = dict(zip(cfg.plan_months, list(cfg.plan_months[1:]) + [cfg.plan_months[-1]]))
        sc = P.set_index(["month", "sku", "chan"]).scale
        P["rec_eom"] = [e * sc[(nxt[m], k, c)] for e, m, k, c in zip(P.eom_target, P.month, P.sku, P.chan)]
        P = P.sort_values(["sku", "chan", "month"]).reset_index(drop=True)
        inv, rb, ro = {}, [], []
        for x in P.itertuples():
            b = inv.get((x.sku, x.chan), x.bom if x.month == cfg.plan_months[0] else 0)
            o = max(x.rec_sales + x.rec_eom - b, 0)
            rb.append(b); ro.append(o); inv[(x.sku, x.chan)] = b + o - x.rec_sales
        P["rec_bom"] = rb; P["rec_otb"] = ro
        P["td"] = [TD[(m, c, cat)] for m, c, cat in zip(P.month, P.chan, P.category)]
        return P


# =============================================================================
# 8. SKU LINE REVIEW FLAGS
# =============================================================================
def review_flags(model: DemandModel, P: pd.DataFrame) -> pd.DataFrame:
    G, L, M = model.grid, model.L, model.M
    jan = P[P.month == CFG.plan_months[0]].copy()
    last12 = G.loc[:, "2025-01-01":"2025-12-01"]
    jan["avg12"] = [last12.loc[(k, c)].mean() for k, c in zip(jan.sku, jan.chan)]
    jan["ly"] = [G.loc[(k, c), pd.Timestamp("2025-01-01")] for k, c in zip(jan.sku, jan.chan)]
    mx = L[(~L.tender) & (L.month >= "2025-01-01")].groupby(["sku", "chan"]).value_k.max()
    jan["maxline"] = [mx.get((k, c), 0) for k, c in zip(jan.sku, jan.chan)]
    jan["conc"] = jan.maxline / np.maximum(jan.avg12 * 12, 1e-9)
    jan["wos"] = jan.bom / np.maximum(jan.sales_plan, 1e-9) * 4.33
    prev = dict(zip(G.index, model.fc11(M - 1, 2)))               # last cycle's view of January
    jan["prev_fc"] = [prev[(k, c)] for k, c in zip(jan.sku, jan.chan)]

    def flags(x):
        f, note = [], []
        eol = x.fading and (x.avg12 > 300 or x.bom > 1000)
        if eol:
            f.append("End-of-life"); note.append("Demand <50% of a year ago vs channel trend: no stock target, buy only to cover the month, then exit")
        if x.bom < x.sales_plan * 0.5 and x.sales_plan > 500 and not eol:
            f.append("Stockout risk"); note.append("Opening stock covers under half of Jan; place PO first")
        if x.wos > 16 and x.bom > 2000:
            f.append("Overstock"); note.append((f"{x.wos:.0f}" if x.wos < 100 else "99+") + " weeks of supply on hand; buy held at zero")
        if x.conc > 0.35 and x.avg12 > 1000:
            f.append("Lumpy"); note.append(f"One line = {x.conc*100:.0f}% of 2025 demand; forecast is fragile")
        if x.ly > 0 and x.sales_plan > 2.5 * x.ly and x.sales_plan > 1000:
            f.append("Forecast > 2.5x LY"); note.append("Check: level inflated by a 2025 spike?")
        if x.prev_fc > 0 and abs(x.sales_plan / x.prev_fc - 1) > 0.3 and x.sales_plan > 1000:
            f.append("Moved >30% vs last cycle")
            note.append(f'Forecast {"up" if x.sales_plan > x.prev_fc else "down"} {abs(x.sales_plan/x.prev_fc-1)*100:.0f}% after Dec actuals')
        return pd.Series({"flags": ", ".join(f) if f else "—", "review_note": "; ".join(note) if note else "Pass"})

    jan = pd.concat([jan, jan.apply(flags, axis=1)], axis=1)
    return jan.sort_values("rec_otb", ascending=False)


# =============================================================================
# 9. LAUNCH-AND-FADE + LIFECYCLE
# =============================================================================
def launch_curve(L: pd.DataFrame, max_age: int = 18) -> pd.DataFrame:
    a = L[~L.sku.str.endswith("99")].copy()
    a["key"] = (a.item_name.fillna("").str.lower().str.replace(r"[^a-z ]", " ", regex=True)
                .str.split().str[:4].str.join(" "))
    a = a[a.key.str.len() > 5]
    first = a.groupby("key").month.min(); n = a.groupby("key").size()
    ok = first[(first >= "2022-10-01") & (first <= "2024-06-01")].index.intersection(n[n >= 10].index)
    b = a[a.key.isin(ok)].copy()
    fm = b.key.map(first)
    b["age"] = (b.month.dt.year - fm.dt.year) * 12 + (b.month.dt.month - fm.dt.month)
    c = b[b.age <= max_age].groupby(["key", "age"]).size().unstack(fill_value=0)
    sh = c.div(c.sum(axis=1), axis=0)
    lc = sh.mean()
    return pd.DataFrame({"month_since_launch": lc.index, "share_pct": (lc.values * 100).round(1),
                         "cum_pct": (lc.cumsum().values * 100).round(1), "items": len(sh)})


def sku_lifecycle(L: pd.DataFrame, name_of, cat_of) -> pd.DataFrame:
    L = L.copy()
    L["half"] = L.month.dt.year.astype(str) + np.where(L.month.dt.month <= 6, "-H1", "-H2")
    hv = L[L.half >= "2022-H2"].groupby(["sku", "half"]).size().unstack(fill_value=0)
    out = []
    for k, row in hv.iterrows():
        last = row.iloc[-2:].sum(); prev = row.iloc[-4:-2].sum()
        ratio = last / prev if prev else np.nan
        st = "Fading" if ratio < 0.75 else ("Growing" if ratio > 1.25 else "Stable")
        rec = {"sku": k, "sku_name": name_of[k], "category": cat_of[k], "peak_half": row.idxmax(),
               "yoy_ratio": round(float(ratio), 2), "status": st}
        rec.update({"h" + h.replace("-", "_"): int(v) for h, v in row.items()})
        out.append(rec)
    return pd.DataFrame(out)


# =============================================================================
# METRICS
# =============================================================================
def wape_bias(x: pd.DataFrame) -> tuple[float, float]:
    return ((x.fc - x.act).abs().sum() / x.act.sum(), (x.fc.sum() - x.act.sum()) / x.act.sum())


VERSION_META = [
    ("v1.0", "Baseline: trailing 3-month average", "Retired", "Highest error; chases the last spike", "2025-07"),
    ("v1.1", "12-month level x category seasonality (shrunk 40% to flat) + tender carve-out", "Live", "Error down 17 pts with bias near zero", "2025-08"),
    ("v1.2a", "v1.1 + SKU trend factor (last 6 mo vs a year earlier)", "Rolled back", "Over-forecasts Bid by 13%; trend picked up the Q4-2024 one-off surge", "2025-09"),
    ("v1.2b", "v1.1 + channel-level trend factor (half strength)", "Rejected", "Bias +14%; same Q4-2024 echo", "2025-10"),
    ("v1.2c", "v1.1 with outlier months capped at 3x the series median", "On hold", "Error down 13 pts but bias -26% fails the gate (|bias| must stay within 10%)", "2025-11"),
]


def version_table(bt: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for ver, change, decision, why, date in VERSION_META:
        x = bt[bt.ver == ver]
        w, b = wape_bias(x)
        wq, _ = wape_bias(x[x.chan == "Quick-Buy"]); wb, _ = wape_bias(x[x.chan == "Bid"])
        rows.append(dict(version=ver, change=change, wape_pct=round(w * 100, 1), bias_pct=round(b * 100, 1),
                         wape_qb_pct=round(wq * 100, 1), wape_bid_pct=round(wb * 100, 1),
                         decision=decision, why=why, date=date))
    V = pd.DataFrame(rows)
    live = V.loc[V.version == "v1.1", "wape_pct"].iloc[0]
    base = V.loc[V.version == "v1.0", "wape_pct"].iloc[0]
    ref = np.where(V.version == "v1.1", base, live)
    V["passes_gate"] = np.where(V.version == "v1.0", "Baseline",
                                np.where((ref - V.wape_pct >= CFG.gate_wape_pts) & (V.bias_pct.abs() <= CFG.gate_bias_abs), "Yes", "No"))
    return V


# =============================================================================
# 10a. CHARTS
# =============================================================================
def _fmt_m(ax):
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))


def chart_otb_raw_vs_rec(plan_rows: pd.DataFrame, path: str):
    g = plan_rows.groupby(["month", "channel"])[["otb_cost_php_m", "rec_otb_cost_php_m"]].sum().reset_index()
    lab = [f"{pd.Timestamp(m + '-01'):%b} {c}" for m, c in zip(g.month, g.channel)]
    x = np.arange(len(g)); w = 0.38
    fig, ax = plt.subplots(figsize=(10, 4.2))
    b1 = ax.bar(x - w / 2 - 0.01, g.otb_cost_php_m, w, color=C_MUTED, label="Model raw")
    b2 = ax.bar(x + w / 2 + 0.01, g.rec_otb_cost_php_m, w, color=C_ORANGE, label="Recommended")
    for bars in (b1, b2):
        ax.bar_label(bars, fmt="%.0f", padding=2, fontsize=8, color=C_TEXT2)
    ax.set_xticks(x, lab); ax.set_ylabel("₱M at cost"); _fmt_m(ax); ax.grid(axis="x", visible=False)
    ax.set_title("Model raw vs recommended buy, Jan–Mar 2026 (₱M at cost)", loc="left")
    ax.legend(loc="upper left", ncol=2)
    fig.savefig(path); plt.close(fig)


def chart_recon(plan_rows: pd.DataFrame, path: str):
    j = plan_rows[plan_rows.month == CFG.plan_months[0]].copy()
    j["label"] = j.channel + " · " + j.category
    j = j.sort_values(["channel", "sales_plan_php_m"])
    y = np.arange(len(j)); h = 0.27
    fig, ax = plt.subplots(figsize=(10, 7))
    ax.barh(y + h, j.sales_plan_php_m, h, color=C_MUTED, label="Model (bottom-up) sales plan")
    ax.barh(y, j.topdown_sales_php_m, h, color=C_BLUE, label="Top-down plan")
    ax.barh(y - h, j.rec_sales_php_m, h, color=C_ORANGE, label="Recommended sales")
    ax.set_yticks(y, j.label); ax.grid(axis="y", visible=False); ax.set_xlabel("₱M (line value)")
    xmax = max(j.sales_plan_php_m.max(), j.topdown_sales_php_m.max())
    ax.set_xlim(0, xmax * 1.38)
    for yi, (v, o) in enumerate(zip(j.variance_pct, j.override)):
        if o != "None":
            ax.text(xmax * 1.36, yi, f"{v:+.0f}% vs TD → scaled", ha="right", va="center", fontsize=8, color=C_TEXT2)
    ax.set_title("January 2026: bottom-up vs top-down by channel × category", loc="left", pad=28)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=3)
    fig.savefig(path); plt.close(fig)


def chart_monthly_demand(md: pd.DataFrame, path: str):
    p = md.pivot_table(index="month", columns="channel", values="base_php_m", aggfunc="sum").sort_index()
    t = md.groupby("month").tender_php_m.sum().reindex(p.index)
    x = pd.to_datetime(p.index + "-01")
    fig, ax = plt.subplots(figsize=(10, 4.2))
    ax.plot(x, p["Bid"], color=C_BLUE, lw=2, label="Bid (base)")
    ax.plot(x, p["Quick-Buy"], color=C_ORANGE, lw=2, label="Quick-Buy (base)")
    ax.set_ylabel("₱M per month"); _fmt_m(ax)
    pk = p["Bid"].idxmax()
    ax.annotate("Oct–Nov 2024 budget release", xy=(pd.Timestamp(pk + "-01"), p["Bid"].max()),
                xytext=(18, -4), textcoords="offset points", fontsize=8, color=C_TEXT2, va="center",
                arrowprops=dict(arrowstyle="-", color=C_TEXT2, lw=0.8))
    ax.set_title("Monthly base demand by channel, ₱M (tender-size orders removed)", loc="left")
    ax.legend(loc="upper left", ncol=2)
    ax.text(0, -0.16, f"Tender-size orders (planned separately) total ₱{t.sum():,.0f}M over the window.",
            transform=ax.transAxes, fontsize=8, color=C_TEXT2)
    fig.savefig(path); plt.close(fig)


def chart_lead_time(lt: pd.DataFrame, path: str):
    fig, ax = plt.subplots(figsize=(8, 2.6))
    for i, r in enumerate(lt.sort_values("median").itertuples()):
        col = C_ORANGE if r.channel == "Quick-Buy" else C_BLUE
        ax.plot([r.p25, r.p75], [i, i], color=col, lw=10, solid_capstyle="round", alpha=0.35)
        ax.plot(r.median, i, "o", color=col, ms=9)
        ax.text(r.p75 + 2, i, f"median {r.median:.0f} days (IQR {r.p25:.0f}–{r.p75:.0f}), n={r.n:,}",
                va="center", fontsize=9, color=C_TEXT)
    ax.set_yticks(range(len(lt)), lt.sort_values("median").channel)
    ax.set_xlim(0, lt.p75.max() + 70); ax.set_ylim(-0.6, len(lt) - 0.4); ax.grid(axis="y", visible=False)
    ax.set_xlabel("Days from request to notice-to-proceed")
    ax.set_title("Lead time by channel", loc="left")
    fig.savefig(path); plt.close(fig)


def chart_seasonality(si: pd.DataFrame, path: str):
    p = si.pivot(index="category", columns="moy", values="index")
    cmap = LinearSegmentedColormap.from_list("div", ["#257de0", "#efefef", "#dc3c3d"])
    fig, ax = plt.subplots(figsize=(11, 4.8))
    im = ax.imshow(p.values, cmap=cmap, norm=TwoSlopeNorm(vmin=0.45, vcenter=1.0, vmax=1.9), aspect="auto")
    ax.set_xticks(range(12), ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    ax.set_yticks(range(len(p)), p.index); ax.grid(False)
    for i in range(p.shape[0]):
        for jj in range(12):
            ax.text(jj, i, f"{p.values[i, jj]:.2f}", ha="center", va="center", fontsize=8, color=C_TEXT)
    for s in ax.spines.values():
        s.set_visible(False)
    cb = fig.colorbar(im, ax=ax, fraction=0.025, pad=0.01); cb.outline.set_visible(False)
    cb.set_label("Index (1.00 = average month)")
    ax.set_title("Seasonality index by category (shrunk 40% toward flat) — v1.1 multipliers", loc="left")
    fig.savefig(path); plt.close(fig)


def chart_launch_curve(lc: pd.DataFrame, path: str):
    fig, ax = plt.subplots(figsize=(10, 3.8))
    cols = [C_BLUE] + [C_MUTED] * (len(lc) - 1)
    bars = ax.bar(lc.month_since_launch, lc.share_pct, color=cols, width=0.75)
    ax.text(lc.month_since_launch.iloc[0], lc.share_pct.iloc[0] + 0.6, f"{lc.share_pct.iloc[0]:.0f}%", ha="center", fontsize=9)
    ax.set_xticks(lc.month_since_launch, [f"M{m}" for m in lc.month_since_launch])
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.grid(axis="x", visible=False); ax.set_xlabel("Months since first appearance")
    ax.set_title(f"Launch-and-fade: share of first-18-month orders by month since launch ({lc['items'].iloc[0]} items)", loc="left")
    m2 = lc.cum_pct.iloc[2]
    ax.text(0.98, 0.9, f"{lc.share_pct.iloc[0]:.0f}% in launch month · {m2:.0f}% by month 2 · ~3% monthly tail, rebuy bump at M12–13",
            transform=ax.transAxes, ha="right", fontsize=8, color=C_TEXT2)
    fig.savefig(path); plt.close(fig)


def chart_lifecycle(lif: pd.DataFrame, path: str):
    halves = [c for c in lif.columns if c.startswith("h20")]
    order = {"Fading": 0, "Growing": 1, "Stable": 2}
    lif = lif.assign(o=lif.status.map(order)).sort_values(["o", "yoy_ratio"])
    n = len(lif); ncol = 7; nrow = int(np.ceil(n / ncol))
    fig, axes = plt.subplots(nrow, ncol, figsize=(14, 1.55 * nrow), sharex=True)
    for ax, r in zip(axes.flat, lif.itertuples()):
        vals = [getattr(r, h) for h in halves]
        col = C_BAD if r.status == "Fading" else (C_OK if r.status == "Growing" else C_BLUE)
        ax.plot(range(len(vals)), vals, color=col, lw=2)
        ax.set_title(f"{r.sku} · {r.status} ({r.yoy_ratio:.2f})", fontsize=8, loc="left", fontweight="normal")
        ax.tick_params(labelsize=6); ax.set_ylim(bottom=0); ax.grid(alpha=0.5)
    for ax in list(axes.flat)[n:]:
        ax.axis("off")
    for ax in axes[-1]:
        ax.set_xticks(range(len(halves)), [h[3:].replace("_", "-") for h in halves], rotation=90, fontsize=6)
    fig.suptitle("SKU lifecycle: order lines per half-year, 2022-H2 → 2025-H2 (ratio = last yr ÷ prior yr)",
                 x=0.01, ha="left", fontsize=12, fontweight="bold")
    fig.tight_layout()
    fig.savefig(path); plt.close(fig)


def chart_versions(V: pd.DataFrame, path: str):
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(11, 3.6))
    cols = [C_BLUE if d == "Live" else C_MUTED for d in V.decision]
    b = a1.bar(V.version, V.wape_pct, color=cols); a1.bar_label(b, fmt="%.1f", fontsize=8, padding=2)
    a1.set_title("Back-test WAPE % (lower is better)", loc="left"); a1.grid(axis="x", visible=False)
    bcol = [C_BAD if abs(v) > CFG.gate_bias_abs else C_OK for v in V.bias_pct]
    b = a2.bar(V.version, V.bias_pct, color=bcol); a2.bar_label(b, fmt="%+.1f", fontsize=8, padding=2)
    a2.axhspan(-CFG.gate_bias_abs, CFG.gate_bias_abs, color=C_GRID, alpha=0.6, zorder=0)
    a2.axhline(0, color=C_TEXT2, lw=0.8)
    a2.set_title("Bias % (gate: within ±10%, shaded)", loc="left"); a2.grid(axis="x", visible=False)
    fig.suptitle("Model versions: v1.1 live; red bias bars fail the ship gate", x=0.01, ha="left", fontweight="bold")
    fig.tight_layout(); fig.savefig(path); plt.close(fig)


def chart_backtest(bm: pd.DataFrame, path: str):
    x = pd.to_datetime(bm.month + "-01")
    fig, ax = plt.subplots(figsize=(10, 4.2))
    ax.plot(x, bm.actual_php_m, color=C_TEXT, lw=2.5, label="Actual")
    ax.plot(x, bm.v1_0_php_m, color=C_AQUA, lw=1.5, label="v1.0 (retired)")
    ax.plot(x, bm.v1_1_php_m, color=C_BLUE, lw=2, label="v1.1 (live)")
    ax.plot(x, bm.v1_2c_php_m, color=C_ORANGE, lw=1.5, label="v1.2c (on hold)")
    ax.set_ylabel("₱M"); _fmt_m(ax)
    ax.set_title("Back-test: actual vs 1-month-ahead forecast, all SKUs (₱M)", loc="left")
    ax.legend(ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.1))
    fig.savefig(path); plt.close(fig)


def chart_inventory_sim(sim: pd.DataFrame, path: str):
    x = np.arange(len(sim)); w = 0.38
    fig, ax = plt.subplots(figsize=(10, 4.2))
    ax.bar(x - w / 2 - 0.01, sim.sales_plan, w, color=C_MUTED, label="Planned sales")
    ax.bar(x + w / 2 + 0.01, sim.act, w, color=C_BLUE, label="Actual sales")
    ax.plot(x, sim.eom, color=C_ORANGE, lw=2, marker="o", ms=6, label="Ending stock")
    ax.set_xticks(x, sim.month); ax.set_ylabel("₱M"); _fmt_m(ax); ax.grid(axis="x", visible=False)
    bias = sim.sales_plan.sum() / sim.act.sum() - 1
    ax.set_title(f"Buys following v1.1, Jul–Dec 2025: forecast ran {bias*100:.0f}% above actual (₱M)", loc="left")
    ax.legend(ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.1))
    fig.savefig(path); plt.close(fig)


def chart_review_flags(jan: pd.DataFrame, path: str):
    cats = ["Stockout risk", "Overstock", "End-of-life", "Lumpy", "Forecast > 2.5x LY", "Moved >30% vs last cycle"]
    data = {c: [jan[(jan.chan == ch) & jan["flags"].str.contains(re.escape(c))].shape[0] for ch in ("Bid", "Quick-Buy")] for c in cats}
    y = np.arange(len(cats)); h = 0.38
    fig, ax = plt.subplots(figsize=(9, 3.6))
    b1 = ax.barh(y + h / 2, [data[c][0] for c in cats], h, color=C_BLUE, label="Bid")
    b2 = ax.barh(y - h / 2, [data[c][1] for c in cats], h, color=C_ORANGE, label="Quick-Buy")
    ax.bar_label(b1, labels=[str(v) if v else "" for v in b1.datavalues.astype(int)], padding=2, fontsize=8)
    ax.bar_label(b2, labels=[str(v) if v else "" for v in b2.datavalues.astype(int)], padding=2, fontsize=8)
    ax.set_yticks(y, cats); ax.invert_yaxis(); ax.grid(axis="y", visible=False)
    ax.set_xlabel("SKU × channel lines"); ax.legend(loc="lower right")
    ax.set_title("January 2026 SKU line review: flags by channel", loc="left")
    fig.savefig(path); plt.close(fig)


# =============================================================================
# 10b. EXCEL PLANNER (live formulas)
# =============================================================================
def build_workbook(P: pd.DataFrame, jan: pd.DataFrame, TD: dict, model: DemandModel, V: pd.DataFrame, path: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    F = "Arial"
    BLUE = Font(name=F, color="0000FF"); BLK = Font(name=F); B = Font(name=F, bold=True)
    H = Font(name=F, bold=True, color="FFFFFF"); GRN = Font(name=F, color="008000")
    HF = PatternFill("solid", fgColor="1F3A5F"); YEL = PatternFill("solid", fgColor="FFFF00")
    wb = Workbook()
    r = wb.active; r.title = "README"
    for i, (t, f) in enumerate([
        ("OTB & Replenishment Planner — January 2026 cycle", B), ("", BLK),
        ("Blue = inputs exported from otb_portfolio.py · Yellow = levers (Assumptions) · Black = formulas · Green = links to another tab.", BLK),
        ("OTB = planned sales + planned EOM stock − opening stock (floored at 0; on-order 0 at cycle start).", BLK),
        ("EOM target = next-month forecast × lead time (months) + safety stock; 0 for end-of-life SKUs.", BLK),
        ("Safety stock = MIN(z × sigma × SQRT(lead time), cap × next-month forecast).", BLK),
        ("Override: cell bottom-up > trigger × top-down → scale its SKUs to target × top-down.", BLK),
        ("Units: PHP thousands (₱k) at line-budget value; cost columns apply the cost ratio.", BLK)], 1):
        r.cell(i, 1, t).font = f
    r.column_dimensions["A"].width = 120
    a = wb.create_sheet("Assumptions")
    A = [("Service level z", CFG.z, "90% cycle service level"),
         ("Lead time Quick-Buy (days)", CFG.lead_days["Quick-Buy"], "Median request→NTP from data"),
         ("Lead time Bid (days)", CFG.lead_days["Bid"], "Median request→NTP from data"),
         ("Safety-stock cap Quick-Buy (months)", CFG.ss_cap_months["Quick-Buy"], ""),
         ("Safety-stock cap Bid (months)", CFG.ss_cap_months["Bid"], "Bid demand visible ~7 weeks ahead"),
         ("Override trigger (BU ÷ TD)", CFG.override_trigger, ""), ("Override target (× TD)", CFG.override_target, ""),
         ("Cost-to-value ratio", CFG.cost_ratio, "Assumption"), ("Overstock threshold (weeks of supply)", 16, ""),
         ("Stockout threshold (BOM ÷ sales)", 0.5, ""), ("Materiality (₱k)", 500, ""), ("Weeks per month", 4.33, "")]
    for j, hh in enumerate(["Lever", "Value", "Note"], 1):
        c = a.cell(1, j, hh); c.font = H; c.fill = HF
    for i, (l, v, n) in enumerate(A, 2):
        a.cell(i, 1, l).font = BLK; c = a.cell(i, 2, v); c.font = BLUE; c.fill = YEL; a.cell(i, 3, n).font = BLK
    a.column_dimensions["A"].width = 38; a.column_dimensions["C"].width = 45
    Z, LTQ, LTB, CQ, CB, TRG, TGT, COST, OS, SO, MAT, WPM = [f"Assumptions!$B${i}" for i in range(2, 14)]

    jp = P[P.month == CFG.plan_months[0]].set_index(["sku", "chan"])
    fp = P[P.month == CFG.plan_months[1]].set_index(["sku", "chan"])
    sigd = dict(zip(model.grid.index, model.sig)); fl = dict(zip(zip(jan.sku, jan.chan), jan["flags"]))
    rows = sorted(jan[["sku", "sku_name", "category", "chan"]].values.tolist(), key=lambda x: (x[3], x[2], x[0]))
    s = wb.create_sheet("SKU_Plan")
    hdr = ["SKU", "Name", "Category", "Channel", "End-of-life (1=yes)", "Opening stock ₱k", "Fcst Jan ₱k", "Fcst Feb ₱k",
           "Error sigma ₱k", "Lead time (months)", "Safety stock ₱k", "EOM target ₱k", "Model OTB ₱k", "Scale Jan",
           "Scale Feb", "Rec. sales Jan ₱k", "Rec. EOM ₱k", "Rec. OTB ₱k", "Rec. OTB at cost ₱k", "Weeks of supply",
           "Stockout flag", "Overstock flag", "Model review flags"]
    for j, hh in enumerate(hdr, 1):
        c = s.cell(1, j, hh); c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, (k, nm, cat, ch) in enumerate(rows, 2):
        x = jp.loc[(k, ch)]
        for j, v in enumerate([k, nm, cat, ch, int(bool(x.fading)), float(x.bom), float(x.sales_plan),
                               float(fp.loc[(k, ch)].sales_plan), float(sigd[(k, ch)])], 1):
            s.cell(i, j, v).font = BLUE if j >= 5 else BLK
        f = {10: f'=IF(D{i}="Bid",{LTB},{LTQ})/30',
             11: f'=MIN({Z}*I{i}*SQRT(J{i}),IF(D{i}="Bid",{CB},{CQ})*H{i})',
             12: f'=IF(E{i}=1,0,H{i}*J{i}+K{i})', 13: f'=MAX(G{i}+L{i}-F{i},0)',
             14: f'=INDEX(Category_Recon!$F:$F,MATCH(D{i}&"|"&C{i},Category_Recon!$A:$A,0))',
             15: f'=INDEX(Category_Recon!$J:$J,MATCH(D{i}&"|"&C{i},Category_Recon!$A:$A,0))',
             16: f'=G{i}*N{i}', 17: f'=L{i}*O{i}', 18: f'=MAX(P{i}+Q{i}-F{i},0)', 19: f'=R{i}*{COST}',
             20: f'=IF(P{i}>0,MIN(F{i}/P{i}*{WPM},99),99)',
             21: f'=IF(AND(E{i}=0,G{i}>{MAT},F{i}<{SO}*G{i}),"Stockout risk","")',
             22: f'=IF(AND(T{i}>{OS},F{i}>2000),"Overstock","")'}
        for j, ff in f.items():
            s.cell(i, j, ff).font = GRN if j in (14, 15) else BLK
        s.cell(i, 23, fl[(k, ch)]).font = BLUE
    n = len(rows) + 1; tot = n + 1
    s.cell(tot, 1, "TOTAL").font = B
    for j in [6, 7, 8, 11, 12, 13, 16, 17, 18, 19]:
        col = s.cell(1, j).column_letter; s.cell(tot, j, f"=SUM({col}2:{col}{n})").font = B
    for row in s.iter_rows(min_row=2, max_row=tot, min_col=6, max_col=20):
        for c in row:
            c.number_format = "#,##0;(#,##0);-"
    for i in range(2, tot + 1):
        for j, fmt in ((10, "0.00"), (14, "0.00"), (15, "0.00"), (20, "0.0")):
            s.cell(i, j).number_format = fmt
    for j, w in enumerate([8, 30, 20, 11, 10, 12, 11, 11, 11, 10, 11, 11, 11, 8, 8, 12, 11, 11, 12, 9, 13, 11, 40], 1):
        s.column_dimensions[s.cell(1, j).column_letter].width = w
    s.freeze_panes = "E2"; s.row_dimensions[1].height = 45

    cr = wb.create_sheet("Category_Recon")
    for j, hh in enumerate(["Key", "Channel", "Category", "BU sales Jan ₱k", "TD plan Jan ₱k", "Scale Jan", "BU/TD Jan",
                            "BU sales Feb ₱k", "TD plan Feb ₱k", "Scale Feb", "Model OTB Jan ₱k", "Rec. OTB Jan ₱k",
                            "Rec. OTB Jan at cost ₱k", "Override?"], 1):
        c = cr.cell(1, j, hh); c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True)
    cells = sorted({(c, cat) for _, _, cat, c in rows})
    for i, (c, cat) in enumerate(cells, 2):
        cr.cell(i, 1, f'=B{i}&"|"&C{i}'); cr.cell(i, 2, c); cr.cell(i, 3, cat)
        sumifs = lambda col: f"=SUMIFS(SKU_Plan!${col}:${col},SKU_Plan!$D:$D,B{i},SKU_Plan!$C:$C,C{i})"
        cr.cell(i, 4, sumifs("G")).font = GRN
        cr.cell(i, 5, float(TD[(CFG.plan_months[0], c, cat)])).font = BLUE
        cr.cell(i, 6, f"=IF(D{i}>{TRG}*E{i},MIN(1,{TGT}*E{i}/D{i}),1)")
        cr.cell(i, 7, f"=IF(E{i}>0,D{i}/E{i},0)")
        cr.cell(i, 8, sumifs("H")).font = GRN
        cr.cell(i, 9, float(TD[(CFG.plan_months[1], c, cat)])).font = BLUE
        cr.cell(i, 10, f"=IF(H{i}>{TRG}*I{i},MIN(1,{TGT}*I{i}/H{i}),1)")
        cr.cell(i, 11, sumifs("M")).font = GRN; cr.cell(i, 12, sumifs("R")).font = GRN
        cr.cell(i, 13, f"=L{i}*{COST}")
        cr.cell(i, 14, f'=IF(F{i}<1,"Scaled to "&TEXT({TGT},"0.00")&"x TD","None")')
        for j in (4, 5, 8, 9, 11, 12, 13):
            cr.cell(i, j).number_format = "#,##0;(#,##0);-"
        for j in (6, 7, 10):
            cr.cell(i, j).number_format = "0.00"
    m = len(cells) + 1; t = m + 1
    cr.cell(t, 2, "TOTAL").font = B
    for j in (4, 5, 8, 9, 11, 12, 13):
        L_ = cr.cell(1, j).column_letter; c = cr.cell(t, j, f"=SUM({L_}2:{L_}{m})"); c.font = B; c.number_format = "#,##0"
    for j, w in enumerate([24, 11, 22, 13, 13, 9, 9, 13, 13, 9, 14, 14, 15, 20], 1):
        cr.column_dimensions[cr.cell(1, j).column_letter].width = w

    v = wb.create_sheet("Version_Log")
    for j, hh in enumerate(["Version", "Tested", "Change", "WAPE %", "Bias %", "WAPE Quick-Buy %", "WAPE Bid %", "Passes gate?", "Decision", "Why"], 1):
        c = v.cell(1, j, hh); c.font = H; c.fill = HF
    for i, x in enumerate(V.itertuples(), 2):
        for j, val in enumerate([x.version, x.date, x.change, x.wape_pct / 100, x.bias_pct / 100, x.wape_qb_pct / 100, x.wape_bid_pct / 100], 1):
            c = v.cell(i, j, val); c.font = BLUE if j >= 4 else BLK
            if j >= 4:
                c.number_format = "0.0%;(0.0%);-"
        v.cell(i, 8, f'=IF(A{i}="v1.0","Baseline",IF(A{i}="v1.1",IF(AND($D$2-D{i}>=0.02,ABS(E{i})<=0.1),"Yes","No"),IF(AND($D$3-D{i}>=0.02,ABS(E{i})<=0.1),"Yes","No")))')
        v.cell(i, 9, x.decision); v.cell(i, 10, x.why)
    for j, w in enumerate([8, 9, 60, 9, 9, 14, 11, 12, 12, 60], 1):
        v.column_dimensions[v.cell(1, j).column_letter].width = w
    wb.save(path)


# =============================================================================
# MAIN
# =============================================================================
def r1(x, n=1):
    return float(round(float(x), n))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--zip", help="Path to COMBINED_PHILGEPS_DEPED.zip")
    ap.add_argument("--csv-dir", help="Or: folder that already holds the extracted CSVs")
    ap.add_argument("--out", default="output", help="Output folder")
    ap.add_argument("--no-excel", action="store_true")
    args = ap.parse_args()
    if not (args.zip or args.csv_dir):
        ap.error("give --zip or --csv-dir")
    out = args.out; charts = os.path.join(out, "charts"); data = os.path.join(out, "data")
    os.makedirs(charts, exist_ok=True); os.makedirs(data, exist_ok=True)
    setup_matplotlib()

    print("[1/9] Loading and cleaning ...")
    folder = args.csv_dir or extract_zip(args.zip)
    raw = load_raw(folder)
    clean_df, dq = clean(raw)
    print(f"      raw {dq['raw_rows']:,} -> clean {dq['clean_rows']:,} rows ({dq['exact_duplicates_removed']:,} duplicates removed)")

    print("[2/9] Mapping SKUs and channels ...")
    lines = map_skus(clean_df)

    print("[3/9] Building demand grid and tender carve-out ...")
    model = DemandModel(lines)
    L = model.L

    print("[4/9] Back-testing model versions (this takes a minute) ...")
    bt = model.backtest()
    V = version_table(bt)
    print(V[["version", "wape_pct", "bias_pct", "decision", "passes_gate"]].to_string(index=False))

    print("[5/9] Simulating Jul–Dec 2025 inventory and planning Jan–Mar 2026 ...")
    sim, bom = model.simulate_inventory()
    P = model.plan(bom)
    TD = model.topdown(P)
    P = model.apply_override(P, TD)

    print("[6/9] Reconciling and reviewing SKU lines ...")
    plan_rows = []
    for (month, c, cat), x in P.groupby(["month", "chan", "category"]):
        bu = x.sales_plan.sum(); td = TD[(month, c, cat)]
        plan_rows.append(dict(month=month, channel=c, category=cat, bom_php_m=r1(x.bom.sum() / 1000, 2),
                              sales_plan_php_m=r1(bu / 1000, 2), eom_target_php_m=r1(x.eom_target.sum() / 1000, 2),
                              otb_retail_php_m=r1(x.otb.sum() / 1000, 2), otb_cost_php_m=r1(x.otb.sum() * CFG.cost_ratio / 1000, 2),
                              rec_sales_php_m=r1(x.rec_sales.sum() / 1000, 2), rec_otb_cost_php_m=r1(x.rec_otb.sum() * CFG.cost_ratio / 1000, 2),
                              rec_bom_php_m=r1(x.rec_bom.sum() / 1000, 2),
                              override="Scaled to 1.25x top-down" if x.scale.min() < 1 else "None",
                              topdown_sales_php_m=r1(td / 1000, 2), variance_pct=r1((bu / td - 1) * 100 if td > 0 else 0, 0),
                              reconcile="Review" if td > 0 and abs(bu / td - 1) > 0.25 else "OK",
                              weeks_of_supply_bom=r1(x.bom.sum() / max(bu, 1e-9) * 4.33, 1)))
    plan_rows = pd.DataFrame(plan_rows)
    jan = review_flags(model, P)

    print("[7/9] Lifecycle, launch curve, demand drivers ...")
    lc = launch_curve(L)
    lif = sku_lifecycle(L, model.name_of, model.cat_of)
    L = L.copy()
    L["lead_days"] = (pd.to_datetime(L.notice_to_proceed_date, errors="coerce") - L.date).dt.days
    ltg = L[(L.lead_days >= 0) & (L.lead_days < 400)].groupby("chan").lead_days
    lead = pd.DataFrame([{"channel": c, "p25": g.quantile(.25), "median": g.median(), "p75": g.quantile(.75), "n": len(g)} for c, g in ltg])
    md = (L.assign(base=np.where(L.tender, 0, L.value_k), tend=np.where(L.tender, L.value_k, 0))
          .groupby(["month", "chan"]).agg(base=("base", "sum"), tender=("tend", "sum"), lines=("value_k", "size")).reset_index())
    md = pd.DataFrame({"month": md.month.dt.strftime("%Y-%m"), "channel": md.chan, "base_php_m": (md.base / 1000).round(2),
                       "tender_php_m": (md.tender / 1000).round(2), "lines": md.lines})
    si = model.si_table(model.M)
    seas = pd.DataFrame([{"category": c, "moy": m, "month": pd.Timestamp(2025, m, 1).strftime("%b"), "index": round(si.loc[m, c], 2)}
                         for c in si.columns for m in range(1, 13)])
    h1 = bt[bt.h == 1].groupby(["target", "ver"])[["fc", "act"]].sum().reset_index()
    bm = h1.pivot(index="target", columns="ver", values="fc") / 1000
    bm.columns = [c.replace(".", "_") + "_php_m" for c in bm.columns]
    bm.insert(0, "actual_php_m", h1.groupby("target").act.first() / 1000)
    bm = bm.reset_index().rename(columns={"target": "month"}).round(2)
    simdf = pd.DataFrame(sim); simdf[simdf.columns[1:]] = (simdf[simdf.columns[1:]] / 1000).round(2)
    tenders = L[L.tender].nlargest(15, "value_k")
    tenders = pd.DataFrame({"month": tenders.month.dt.strftime("%Y-%m"), "sku": tenders.sku, "channel": tenders.chan,
                            "entity": tenders.procuring_entity.str[:60],
                            "item": tenders.item_name.fillna(tenders.notice_title).str[:70], "php_m": (tenders.value_k / 1000).round(1)})
    dq_log = pd.DataFrame([
        ("Raw rows loaded (19 quarterly files)", dq["raw_rows"], "Combined all files"),
        ("Rows sitting in the wrong quarter file", dq["rows_outside_file_quarter"], "Kept; re-keyed on the record date"),
        ("Duplicate rows removed (1.0 vs 1, True vs TRUE)", dq["exact_duplicates_removed"], "Normalized formats, deduped on 10 fields"),
        ("2022-Q1 file missing", dq["rows_by_quarter"].get("2022-Q1", 0), "Modeling window starts Apr-2022"),
        ("Lines with no item budget", int(round(dq["missing_item_budget_share"] * dq["clean_rows"])), "Notice budget split across its lines"),
        ("Service lines excluded", int(dq["clean_rows"] - len(L)), "Out of goods replenishment scope"),
        ("Goods lines mapped to 39 SKUs, Apr-2022 to Dec-2025", len(L), "Keyword rules, then mixed lots by category"),
        ("Tender-size lines carved out (> channel P99)", int(L.tender.sum()), "Planned as a separate pipeline")],
        columns=["step", "rows", "action"])

    S = simdf
    kpis = {"clean_lines": len(L), "skus": int(model.grid.index.get_level_values(0).nunique()),
            "jan_otb_cost_php_m": r1(jan.otb.sum() * CFG.cost_ratio / 1000),
            "jan_rec_otb_cost_php_m": r1(jan.rec_otb.sum() * CFG.cost_ratio / 1000),
            "q1_otb_cost_php_m": r1(P.otb.sum() * CFG.cost_ratio / 1000),
            "q1_rec_otb_cost_php_m": r1(P.rec_otb.sum() * CFG.cost_ratio / 1000),
            "champ_wape_pct": float(V.loc[V.version == "v1.1", "wape_pct"].iloc[0]),
            "champ_bias_pct": float(V.loc[V.version == "v1.1", "bias_pct"].iloc[0]),
            "at_risk_skus": int((jan["flags"] != "—").sum()),
            "stockout_risk": int(jan["flags"].str.contains("Stockout").sum()),
            "overstock": int(jan["flags"].str.contains("Overstock").sum()),
            "eol": int(jan["flags"].str.contains("End-of-life").sum()),
            "jan_cells_overridden": int((plan_rows[plan_rows.month == CFG.plan_months[0]].override != "None").sum()),
            "h2_2025_bias_pct": r1((S.sales_plan.sum() / S.act.sum() - 1) * 100, 0),
            "tender_value_share_pct": r1(L[L.tender].value_k.sum() / L.value_k.sum() * 100, 0),
            "launch_m0_pct": r1(lc.share_pct.iloc[0], 0), "launch_items": int(lc["items"].iloc[0])}
    kpis["jan_cut_pct"] = r1((1 - kpis["jan_rec_otb_cost_php_m"] / kpis["jan_otb_cost_php_m"]) * 100, 0)
    kpis["q1_cut_pct"] = r1((1 - kpis["q1_rec_otb_cost_php_m"] / kpis["q1_otb_cost_php_m"]) * 100, 0)

    print("[8/9] Writing datasets and charts ...")
    sku_out = jan[["sku", "sku_name", "category", "chan", "bom", "prev_fc", "sales_plan", "rec_sales", "eom_target",
                   "otb", "rec_otb", "wos", "flags", "review_note"]].rename(columns={"chan": "channel"})
    sku_out["rec_otb_cost"] = sku_out.rec_otb * CFG.cost_ratio
    for name, df in {"otb_plan": plan_rows, "sku_buy_jan": sku_out.round(1), "monthly_demand": md, "seasonality": seas,
                     "lead_time": lead, "launch_curve": lc, "sku_lifecycle": lif, "model_versions": V,
                     "backtest_monthly": bm, "inventory_sim": simdf, "data_quality": dq_log, "tenders": tenders}.items():
        df.to_csv(os.path.join(data, f"{name}.csv"), index=False)
    json.dump(kpis, open(os.path.join(data, "kpis.json"), "w"), indent=1)

    chart_otb_raw_vs_rec(plan_rows, os.path.join(charts, "01_otb_raw_vs_recommended.png"))
    chart_recon(plan_rows, os.path.join(charts, "02_jan_topdown_reconciliation.png"))
    chart_review_flags(jan, os.path.join(charts, "03_sku_review_flags.png"))
    chart_monthly_demand(md, os.path.join(charts, "04_monthly_demand_by_channel.png"))
    chart_lead_time(lead, os.path.join(charts, "05_lead_time_by_channel.png"))
    chart_seasonality(seas, os.path.join(charts, "06_seasonality_heatmap.png"))
    chart_launch_curve(lc, os.path.join(charts, "07_launch_and_fade_curve.png"))
    chart_lifecycle(lif, os.path.join(charts, "08_sku_lifecycle_small_multiples.png"))
    chart_versions(V, os.path.join(charts, "09_model_versions_wape_bias.png"))
    chart_backtest(bm, os.path.join(charts, "10_backtest_actual_vs_forecast.png"))
    chart_inventory_sim(simdf, os.path.join(charts, "11_inventory_simulation_h2_2025.png"))

    if not args.no_excel:
        print("[9/9] Building Excel OTB planner ...")
        build_workbook(P, jan, TD, model, V, os.path.join(out, "OTB_Planner_Jan2026.xlsx"))

    print("\nHeadline KPIs")
    for k, v in kpis.items():
        print(f"  {k:28s} {v}")
    print(f"\nDone. Outputs in {os.path.abspath(out)}")


if __name__ == "__main__":
    main()
