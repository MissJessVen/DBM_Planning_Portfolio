# OTB & Replenishment Planning Portfolio

**An end-to-end open-to-buy (OTB) and replenishment planner, built in Python on 221,000 rows of real, messy, high-SKU procurement data.**

It turns monthly SKU-level demand forecasts into a reconciled replenishment buy, and it does that across two channels with different demand drivers. The project covers the whole workflow: cleaning the data, building a SKU master, forecasting, back-testing every logic change against a ship gate, simulating inventory, reconciling bottom-up to top-down, reviewing SKU lines for outliers, and modeling launch-and-fade lifecycles.

![Model raw vs recommended buy](output/charts/01_otb_raw_vs_recommended.png)

---

## Contents
- [Headline results](#headline-results)
- [Why this dataset](#why-this-dataset)
- [Quick start](#quick-start)
- [Project structure](#project-structure)
- [Methodology](#methodology)
- [Model versions and the ship gate](#model-versions-and-the-ship-gate)
- [Charts](#charts)
- [Excel OTB planner](#excel-otb-planner)
- [Assumptions and limitations](#assumptions-and-limitations)
- [Skills demonstrated](#skills-demonstrated)

---

## Headline results

These are for the January 2026 planning cycle.

| Metric | Value |
|---|---|
| Recommended January buy (at cost) | **₱308.1M**, 35% below the model's raw ₱470.8M |
| Recommended Jan–Mar buy (at cost) | **₱875.5M**, 60% below the model's raw ₱2,173.6M |
| Channel × category cells scaled back after top-down reconciliation | 11 of 18 |
| Live model error (WAPE, SKU × channel × month) | 92.2%, bias −1.8% |
| Improvement vs baseline | −16.9 WAPE points |
| SKU lines flagged for review before purchase orders | 40 of 78 (7 stockout risk, 19 overstock, 17 end-of-life) |
| Share of demand in an item's launch month | 38% (51% by month 2), across 92 new items |

**The main finding:** the live model over-forecast July–December 2025 by 50%. Demand fell after a one-time budget release in Q4 2024. Following the model's buys left too much stock in the Bid channel, which is why the January buy needs a top-down guardrail. That is also why the most promising model change (v1.2c) is still on hold: it lowers error but under-forecasts by 26%.

---

## Why this dataset

This project uses **PhilGEPS** (Philippine Government Electronic Procurement System) purchase requests from the **Department of Education**, 2021–2025. Each request line is treated as a customer order for a supplier who has to buy stock ahead of demand. The mapping keeps every hard part of a real retail planning problem:

| Retail concept | Stand-in used here |
|---|---|
| DTC channel | **Quick-Buy**: small-value and shopping purchases. Many small orders, ~2-week lead time. |
| Amazon / FBA channel | **Bid**: public bidding and other formal modes. Fewer, larger, lumpy orders, ~7-week lead time. |
| SKU | 39 SKUs built from free-text item names with keyword rules, plus "mixed lot" SKUs by business category. |
| Category | 9 categories (Office & Print, IT & Devices, Food & Nutrition, Learning Materials, …) |
| Holiday / gifting peak | Q4 budget-release surge (Oct–Nov), plus school-opening and recognition-season peaks |
| Launch-and-fade | New item names: first appearance, rebuy curve, end-of-life rule |
| Sales value | Line budget in Philippine pesos (₱) |

---

## Quick start

```bash
git clone https://github.com/<your-username>/<repo-name>.git
cd <repo-name>
pip install -r requirements.txt
python otb_portfolio.py --zip data/COMBINED_PHILGEPS_DEPED.zip --out output
```

The run takes about 15 seconds and writes the charts, datasets and Excel planner to `output/`.

| Flag | Purpose |
|---|---|
| `--zip PATH` | The raw zip of the 19 quarterly CSVs |
| `--csv-dir PATH` | Use a folder of already-extracted CSVs instead |
| `--out PATH` | Output folder (default `output`) |
| `--no-excel` | Skip building the Excel planner |

**Requirements:** Python 3.10+, pandas, numpy, matplotlib, openpyxl (see `requirements.txt`).

All planning settings sit in the `Config` class at the top of `otb_portfolio.py`: service level, lead times, safety-stock caps, the override rule, cost ratio and the ship gate. Change one and re-run to see the effect.

---

## Project structure

```
.
├── otb_portfolio.py          # the full pipeline
├── requirements.txt
├── README.md
├── data/
│   └── COMBINED_PHILGEPS_DEPED.zip   # raw PhilGEPS DepEd CSVs (19 quarterly files)
└── output/
    ├── OTB_Planner_Jan2026.xlsx      # formula-driven OTB planner
    ├── charts/                       # 11 PNG charts
    └── data/                         # CSV datasets + kpis.json
```

---

## Methodology

### 1. Data cleaning
The 19 quarterly files did not line up with their names, and the same values were written in different formats.

| Issue found | Rows | Fix |
|---|---:|---|
| Raw rows loaded | 221,302 | — |
| Rows in the wrong quarter's file | 40,355 | Re-keyed on record date, not file name |
| Duplicate rows across files (`3840` vs `3840.0`, `True` vs `TRUE`) | 56,269 | Normalized formats, then de-duplicated on 10 fields |
| 2022-Q1 file missing | — | Modeling window starts Apr-2022; the gap is not treated as zero demand |
| Lines with no item budget | 83,342 | Notice budget split evenly across its lines |
| Service lines (catering, venues, lodging, repairs) | 98,529 | Excluded from goods replenishment |
| Goods lines mapped to 39 SKUs | 66,504 | Modeling base |
| Tender-size lines (above each channel's 99th percentile) | 866 | Planned as a separate pipeline. Only 866 lines, but 40% of value. |

### 2. Forecasting
- **Level:** mean of the last 12 months of base demand per SKU × channel.
- **Seasonality:** a category × month-of-year index, shrunk 40% toward flat so one noisy year can't dominate.
- **Back-test:** 16 rolling origins (Jul-2024 to Oct-2025) at 1–3 month horizons, scored on WAPE and bias.

### 3. OTB engine
```
OTB            = planned sales + planned end-of-month stock − opening stock   (floored at 0)
EOM target     = next-month forecast × lead time (months) + safety stock     (0 for end-of-life SKUs)
Safety stock   = MIN(z × σ_error × √lead time,  cap × next-month forecast)
                 z = 1.28 (90% service); cap = 1.0 month Quick-Buy, 0.5 month Bid
```
- **Lead times** come from the data: median 14 days for Quick-Buy and 51 days for Bid, measured from request to notice-to-proceed.
- **Opening stock** is not assumed. It comes from re-running the planner's own buys for July–December 2025 against what was actually ordered, so it carries the model's real forecast misses.

### 4. Top-down reconciliation and analyst override
- **Top-down plan:** the 2023–2025 same-month average × channel-category momentum (H2-2025 ÷ H2-2024, square-rooted, bounded 0.5–1.5).
- **Override:** if a cell's bottom-up sales plan is more than 1.5× top-down, all its SKUs are scaled down to 1.25× top-down before the buy becomes a purchase order.

### 5. SKU line review
Every January line runs through six checks. Each check fires only above a materiality threshold, so small SKUs don't bury the real issues.

| Flag | Rule |
|---|---|
| End-of-life | Last 6 months below 50% of the same 6 months a year earlier, relative to the channel's own trend |
| Stockout risk | Opening stock covers less than half of January's forecast |
| Overstock | More than 16 weeks of supply on hand |
| Lumpy | A single order is more than 35% of 2025 demand |
| Forecast > 2.5× last year | Possible spike inflating the level |
| Moved > 30% vs last cycle | Explains "what changed since last month" |

### 6. Launch-and-fade
Across 92 items first seen between Oct-2022 and Jun-2024 (10+ orders each), 38% of 18-month demand lands in the launch month and 51% by month 2. After that, demand settles to a ~3% monthly tail, with a bump at months 12–13 from the annual rebuy. That shape sets three rules:
- **New items:** buy the launch month off the first order book, not history.
- **Rebuys:** cap the month-1 rebuy at ~25% of launch volume.
- **End-of-life:** stop rebuying under the end-of-life rule above.

---

## Model versions and the ship gate

A change ships only if it cuts WAPE by **at least 2 points** *and* keeps **|bias| ≤ 10%**. The bias test exists because a model that under-forecasts can show lower error while it starves the buy.

| Version | Change | WAPE | Bias | Decision |
|---|---|---:|---:|---|
| v1.0 | Trailing 3-month average | 109.1% | +3.9% | Retired |
| **v1.1** | **12-month level × category seasonality + tender carve-out** | **92.2%** | **−1.8%** | **Live** |
| v1.2a | + SKU trend factor | 97.0% | +12.2% | Rolled back |
| v1.2b | + channel-level trend factor | 98.5% | +13.9% | Rejected |
| v1.2c | Outlier months capped at 3× median | 79.2% | −26.4% | On hold (fails bias gate) |

**Next test:** flag the Q4-2024 budget release as a one-time event so it drops out of the 12-month level, and pair v1.2c's capped level with a bias correction.

---

## Charts

### Buy and reconciliation
| | |
|---|---|
| ![](output/charts/01_otb_raw_vs_recommended.png) | ![](output/charts/02_jan_topdown_reconciliation.png) |
| ![](output/charts/03_sku_review_flags.png) | ![](output/charts/11_inventory_simulation_h2_2025.png) |

### Demand drivers
| | |
|---|---|
| ![](output/charts/04_monthly_demand_by_channel.png) | ![](output/charts/05_lead_time_by_channel.png) |

![Seasonality heatmap](output/charts/06_seasonality_heatmap.png)

### Lifecycle
![Launch-and-fade curve](output/charts/07_launch_and_fade_curve.png)
![SKU lifecycle](output/charts/08_sku_lifecycle_small_multiples.png)

### Model performance
| | |
|---|---|
| ![](output/charts/09_model_versions_wape_bias.png) | ![](output/charts/10_backtest_actual_vs_forecast.png) |

---

## Excel OTB planner

`output/OTB_Planner_Jan2026.xlsx` is a formula-driven planner whose totals tie exactly to the Python output. It has 1,216 live formulas.

| Tab | What it holds |
|---|---|
| **Assumptions** | Every lever in one place (yellow cells) |
| **SKU_Plan** | 78 SKU × channel lines: opening stock, forecast, safety stock, EOM target, model OTB, override scale, recommended OTB, flags |
| **Category_Recon** | Bottom-up vs top-down by channel × category; computes the override scale |
| **Version_Log** | Back-test results, with the ship gate evaluated by formula |

Color convention: blue = inputs from the pipeline · yellow = planning levers · black = formulas · green = links to another tab.

---

## Assumptions and limitations

- **Proxy data:** procurement requests stand in for customer orders. There are no returns, promotions or site-traffic signals.
- **Values** are line budgets in PHP thousands, not invoiced sales. **Buy cost is assumed at 70% of line value.**
- **Opening stock is simulated** from July–December 2025 buys, not taken from a real inventory system.
- **"Mixed lot" SKUs** (generic lines such as "office supplies") hide item-level detail.
- **The 2022-Q1 source file is missing**, so the modeling window starts in Apr-2022.
- **SKU-level accuracy for Bid is low** (~98% WAPE) because its demand is tender-driven and lumpy. That is why Bid is planned with lower safety-stock caps and a top-down guardrail.

---

## Skills demonstrated

- **Forecasting and statistics:** seasonal indices with shrinkage, rolling-origin back-testing, WAPE and bias, error-based safety stock.
- **Retail planning:** OTB, weeks of supply, EOM targets, top-down / bottom-up reconciliation, end-of-life rules.
- **Short-lifecycle modeling:** launch-and-fade curves and rebuy / stop-rebuy rules.
- **Model governance:** versioned logic changes, a ship gate, documented rollbacks.
- **Data engineering:** de-duplicating across overlapping files, format normalization, rule-based SKU mapping on free text.
- **Communication:** clear charts, a reconciled Excel planner, and an explainable "why is the model saying this" trail for every SKU line.

---

## Data source

PhilGEPS (Philippine Government Electronic Procurement System) open procurement data, Department of Education records, 2021–2025.
