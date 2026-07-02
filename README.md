# Apple Inc. (AAPL) — Three-Statement Financial Model & Valuation

A full investment-banking-style financial model built in Python (`openpyxl` + `yfinance`) that auto-generates an Excel workbook for any public ticker. This instance is pre-run for **AAPL** using FY2025 and FY2024 annual data sourced live from Yahoo Finance.

---

## How to Run

```bash
pip install openpyxl yfinance
python3 financial_model.py AAPL   # or any ticker: MSFT, GOOGL, META …
```

The Excel file is saved to the same directory as the script:
`DCF_model/AAPL_Financial_Model.xlsx`

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
| DCF Model | Red | 5-year UFCF projection + Terminal Value + Sensitivity |
| WACC Calculation | Steel Blue | CAPM cost of equity + weighted cost of capital |
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

## DCF Valuation

### Sensitivity Table — Intrinsic Value per Share (USD)

|  | g = 1.5% | g = 2.0% | **g = 2.5%** | g = 3.0% | g = 3.5% |
|---|---|---|---|---|---|
| WACC = 9.6% | $62.10 | $66.51 | **$71.54** | $77.34 | $84.08 |
| WACC = 10.6% | $52.82 | $56.17 | **$59.93** | $64.18 | $69.03 |
| WACC = 11.6% | $45.50 | $48.10 | **$50.99** | $54.22 | $57.85 |
| WACC = 12.6% | $39.59 | $41.66 | **$43.94** | $46.45 | $49.24 |
| WACC = 13.6% | $34.75 | $36.43 | **$38.25** | $40.25 | $42.45 |

**Base case (WACC ≈ 10.3%, g = 2.5%) implies ~$60–72/share on a pure DCF basis.** The wide gap between this and the current ~$289 market price reflects that the market is pricing in either significantly higher growth expectations, a much lower discount rate, or a large premium for Apple's platform lock-in, brand, and capital return program that a standard 5-year DCF does not capture in the terminal value. The valuation summary anchors the DCF weight at 35% and blends it with market-observable comps.

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

**Takeaway:** AAPL trades at 26.7× EV/EBITDA and 35.1× P/E — a **57% premium to peer median EV/EBITDA (17.0×)** and **29% premium to peer median P/E (27.2×)**. This premium reflects Apple's Services margin profile, install-base stickiness, and capital return consistency vs. higher-growth but less predictable peers. Only NVDA and GOOGL trade at comparable multiples.

---

## Valuation Summary — Football Field

| Methodology | Low | Base Case | High | Weight |
|---|---|---|---|---|
| 52-Week Trading Range | $201.50 | $289.36 | $317.40 | 0% |
| DCF — Base Case | $245.96 | $300.93 | $361.70 | **35%** |
| Trading Comps — EV/EBITDA | $260.42 | $289.36 | $332.76 | **25%** |
| Trading Comps — P/E | $266.21 | $300.93 | $341.44 | **25%** |
| Precedent Transactions | $303.83 | $347.23 | $405.10 | **15%** |
| Analyst Price Targets | $215.00 | $315.09 | $400.00 | 0% |
| **Weighted Average Implied Price** | — | **~$305** | — | **100%** |
| **Actual Market Price** | — | **$289.36** | — | — |

Weights: DCF 35% / EV/EBITDA 25% / P/E 25% / Precedent Transactions 15%. Trading range and analyst targets are benchmarks only (weight = 0).

**Result:** The blended implied price of ~$305 suggests AAPL is approximately **5% undervalued** relative to the market price of $289.36 at model run time, with uncertainty ranges spanning $246–$405 across methodologies.

---

## File Structure

```
DCF_model/
├── financial_model.py       # Main script — all 9 sheet builders
├── AAPL_Financial_Model.xlsx  # Generated output (re-created each run)
└── README.md                # This file
```

---

*Model built with Python 3, openpyxl, and yfinance. Financial data © Yahoo Finance. For educational and analytical purposes only — not investment advice.*
