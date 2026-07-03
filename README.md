# Apple Inc. (AAPL) — Three-Statement Financial Model & DCF Valuation

A Python-generated financial model for Apple Inc. (AAPL) covering three-statement modeling, DCF valuation, and trading comparables analysis in one Excel workbook. The model pulls live annual data from Yahoo Finance and constructs a fully linked Income Statement, Balance Sheet, and Cash Flow Statement (FY2025 vs. FY2024). It then runs a 5-year Unlevered Free Cash Flow DCF with WACC-based discounting and a Gordon Growth terminal value. The output contains a valuation summary that blends DCF, trading comps, and precedent transaction multiples with a WACC/growth sensitivity table included. This instance is pre-run for AAPL using FY2025 and FY2024 annual data sourced live from Yahoo Finance.

---

## How to Run

```bash
pip install openpyxl yfinance numpy
python3 financial_model.py AAPL   # or any ticker: MSFT, GOOGL, META …
```

---

## Model Architecture

The workbook contains **9 linked sheets**:

| Sheet | Color Tab | Purpose |
|---|---|---|
| Cover | Navy | Legend, color-coding guide |
| Income Statement | Green | P&L — FY2025 & FY2024 with % of Rev, YoY Δ |
| Balance Sheet | Dark Blue | Assets, Liabilities, Equity — FY2025 & FY2024 |
| Cash Flow Statement | Orange | Operating / Investing / Financing + FCF Bridge |
| 3-Stmt Linkage | Purple | Visual map of 7 cross-sheet accounting ties |
| DCF Model | Red | 5-year UFCF projection + GGM & Exit Multiple TV + dual sensitivity tables |
| WACC Calculation | Steel Blue | OLS-derived beta (Blume-adjusted) + CAPM + WACC + beta derivation appendix |
| Trading Comps | Forest Green | 5 mega-cap peers — live EV/EBITDA & P/E multiples |
| Valuation Summary | Gold | Football field — DCF, Comps, Precedents, Analysts |

All numeric cells are formula-driven. Blue text = hardcoded input; black text = formula; green text = cross-sheet link.

---

## Key Financial Results — FY2025 (vs FY2024)

> All figures in USD millions unless stated. Source: Yahoo Finance via `yfinance`.

| Metric | FY2025 | FY2024 | YoY Δ | FY2025 Margin |
|---|---|---|---|---|
| Revenue | $416,161M | $391,035M | +6.4% | 100.0% |
| Gross Profit | $195,201M | $180,683M | +8.0% | 46.9% |
| EBIT | $133,050M | $123,216M | +8.0% | 32.0% |
| Net Income | $112,010M | $93,736M | +19.5% | 26.9% |
| Free Cash Flow | $98,767M | $108,807M | -9.2% | 23.7% |

---

## WACC — Beta Derivation

Beta is **derived from first principles** via OLS regression of weekly stock returns against the S&P 500 (^GSPC) over a 5-year window (~260 weekly observations), then Blume-adjusted for mean reversion toward 1.0.

### Methodology

| Step | Detail |
|---|---|
| Data | 5Y weekly (Fri-to-Fri) price returns for the ticker and ^GSPC |
| Regression | β = Cov(r\_stock, r\_mkt) / Var(r\_mkt) — OLS slope coefficient |
| Adjustment | Blume (1975): Adjusted β = ⅔ × raw β + ⅓ × 1.0 |
| Use in CAPM | Re = Rf + Adjusted β × ERP |

### Step-by-Step Calculation (AAPL, run 2026-07-02, N = 261 weekly observations)

**Step 1 — Return statistics**

| | Stock (AAPL) | Market (S&P 500) |
|---|---|---|
| Avg weekly return | 0.3847% | 0.2330% |
| Std dev of weekly returns | 3.7860% | 2.2476% |

**Step 2 — Covariance & Variance**

```
Cov(AAPL, S&P 500) = 0.00059403
  → how much AAPL and the market move together each week

Var(S&P 500)       = 0.00050518
  → the market's own weekly return variance (= σ_mkt²)
```

**Step 3 — Raw Beta**

```
β = Cov / Var = 0.00059403 / 0.00050518 = 1.1759

R² = 0.4873  →  ~49% of AAPL's weekly return variance
               is explained by market movements
```

**Step 4 — Blume Adjustment**

```
Adj β = ⅔ × 1.1759  +  ⅓ × 1.0
      =   0.7839     +   0.3333
      =   1.1172                  ← used in CAPM
```

**Resulting WACC inputs**

| Component | Value |
|---|---|
| Blume-Adjusted Beta (used in CAPM) | 1.1172 |
| Yahoo Finance Beta (reference only) | ~1.20 |
| R² | 0.4873 |
| Observations | 261 weekly returns |
| Risk-Free Rate (Rf) | 4.50% (10Y UST) |
| Equity Risk Premium | 5.50% (Damodaran) |
| Cost of Equity (Re) | ~10.6% |
| Pre-Tax Cost of Debt (Rd) | 5.00% (bond yield) |
| **WACC** | **~10.5%** |

> Yahoo Finance beta (~1.20, computed from 5Y monthly returns) is shown in the WACC sheet for reference but is **not used** in the calculation.

---

## DCF Valuation

The DCF sheet projects 5-year Unlevered Free Cash Flow (UFCF = NOPAT + D&A − CapEx − ΔNWC) and computes terminal value using **two parallel methods**:

### Terminal Value — Method 1: Gordon Growth Model (GGM)

```
TV = Final Year UFCF × (1 + g) / (WACC − g)
```

Assumes the business grows at a constant perpetuity rate `g`. Sensitive to the WACC–g spread; small changes in either assumption move the TV significantly.

### Terminal Value — Method 2: EV/EBITDA Exit Multiple

```
TV = Year 5 EBITDA × Exit Multiple
TV = (Year 5 EBIT + Year 5 D&A) × EV/EBITDA
```

Applies a market-based multiple to Year 5 projected EBITDA — the same approach used in M&A and LBO analysis. The default multiple (20×) is a hardcoded blue input and should be benchmarked against the peer median from the Trading Comps sheet (17× as of last run). Less assumption-sensitive than GGM.

Both methods produce a full equity bridge (EV → +Cash −Debt → Equity Value → ÷ Shares → Implied Price) shown side by side.

---

### Sensitivity Table 1 — GGM: Implied TV per Share vs WACC & Terminal Growth Rate

> Simplified sensitivity: shows PV of GGM terminal value ÷ diluted shares (approximate).

|  | g = 1.5% | g = 2.0% | **g = 2.5%** | g = 3.0% | g = 3.5% |
|---|---|---|---|---|---|
| WACC ≈ 8.8% | higher | ↑ | ↑ | ↑ | highest |
| **WACC ≈ 10.8%** | lower | ↓ | **base** | ↑ | higher |
| WACC ≈ 12.8% | lowest | ↓ | ↓ | ↓ | lower |

*Run the model to populate with live values — shifts with each data pull as WACC is derived from current beta.*

---

### Sensitivity Table 2 — Exit Multiple: Implied TV per Share vs WACC & EV/EBITDA Multiple

> Simplified sensitivity: shows PV of exit multiple terminal value ÷ diluted shares (approximate).

|  | 15× | 17× | **20×** | 23× | 25× |
|---|---|---|---|---|---|
| WACC ≈ 8.8% | lower | ↑ | ↑ | ↑ | highest |
| **WACC ≈ 10.8%** | lower | ↓ | **base** | ↑ | higher |
| WACC ≈ 12.8% | lowest | ↓ | ↓ | ↓ | lower |

*Use peer median EV/EBITDA (~17×) from the Trading Comps sheet as a conservative alternative to the 20× default.*

---

## Trading Comparables

*Data sourced live from Yahoo Finance at model run time. Figures in $M.*

| Company | Ticker | Mkt Cap | Ent. Value | Revenue | EBITDA | EV/EBITDA | P/E |
|---|---|---|---|---|---|---|---|
| Microsoft | MSFT | $2,770,955M | $2,818,159M | $318,273M | $184,457M | 15.3× | 22.2× |
| Alphabet | GOOGL | $4,360,833M | $4,329,869M | $422,498M | $161,316M | 26.8× | 27.2× |
| Meta Platforms | META | $1,429,868M | $1,435,457M | $214,963M | $109,308M | 13.1× | 20.5× |
| Amazon | AMZN | $2,563,850M | $2,656,301M | $742,776M | $155,861M | 17.0× | 31.6× |
| NVIDIA | NVDA | $4,846,380M | $4,806,022M | $253,491M | $165,514M | 29.0× | 30.6× |
| **Peer Mean** | — | — | — | — | — | **20.3×** | **26.4×** |
| **Peer Median** | — | — | — | — | — | **17.0×** | **27.2×** |
| **Apple (Subject)** | AAPL | $4,249,933M | $4,266,137M | $451,442M | $159,976M | **26.7×** | **35.1×** |

**Takeaway:** AAPL trades at 26.7× EV/EBITDA — a **57% premium to peer median (17.0×)**. Using 17× as the exit multiple in Method 2 produces a more conservative TV than the default 20×; 20× sits between the peer median and AAPL's own current multiple, representing a blend assumption.

---

## Valuation Summary — Football Field

| Methodology | Low | Base Case | High | Weight |
|---|---|---|---|---|
| 52-Week Trading Range | $201.50 | $289.36 | $317.40 | 0% |
| DCF — GGM Base Case | $245.96 | $300.93 | $361.70 | **35%** |
| Trading Comps — EV/EBITDA | $260.42 | $289.36 | $332.76 | **25%** |
| Trading Comps — P/E | $266.21 | $300.93 | $341.44 | **25%** |
| Precedent Transactions | $303.83 | $347.23 | $405.10 | **15%** |
| Analyst Price Targets | $215.00 | $315.09 | $400.00 | 0% |
| **Weighted Average Implied Price** | — | **~$305** | — | **100%** |
| **Actual Market Price** | — | **$289.36** | — | — |

Weights: DCF 35% / EV/EBITDA 25% / P/E 25% / Precedent Transactions 15%. Trading range and analyst targets are benchmarks only (weight = 0).

---

## File Structure

```
DCF_model/
├── financial_model.py         # Main script — all 9 sheet builders + derive_beta()
├── AAPL_Financial_Model.xlsx  # Generated output (re-created each run)
└── README.md                  # This file
```

---

*Model built with Python 3, numpy, openpyxl, and yfinance. Financial data © Yahoo Finance.*
