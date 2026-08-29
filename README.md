# Paramount Resources (TSX: POU) — Energy Equity Research & Valuation

An end-to-end equity research case study on **Paramount Resources Ltd.**, a Canadian oil and natural gas producer with a market capitalisation of approximately C$4.8 billion.

The project builds a complete analytical chain — from reconciled financial disclosure, through a bottom-up well-level production forecast, to an asset-level valuation and a written recommendation with a price target.

**Conclusion: HOLD, target C$33.00** against a share price of C$32.65. Three independent valuation methods converge within 9%.

> Hypothetical analytical exercise prepared for illustrative purposes. Not investment advice, and not a recommendation to buy or sell any security. The author holds no position.

---

## The written output

📄 **[`Paramount_Resources_Initiating_Coverage.pdf`](Paramount_Resources_Initiating_Coverage.pdf)** — a 5-page initiating-coverage note. Start here.

Everything else in this repository is the working paper behind it.

---

## Headline findings

**The outspend is construction, not structure.** Paramount has produced negative free cash flow for six consecutive quarters. Decomposing this at the well level shows that sustaining the Willesden Green plateau of 50,000 Boe/d requires roughly 16 wells per year at C$210mm, against C$803mm of annual field netback — maintenance capital of 26% of cash flow. The current cash burn is the Alhambra and Sinclair build-outs, which are finite.

**Volume growth is not value growth.** Production roughly doubles by 2030, but the incremental barrels are Sinclair dry gas. Liquids weighting falls from 49% to 32% and netback per Boe falls from C$38 to C$30. The 6:1 Boe conversion is an *energy* equivalence, not a *value* equivalence.

**The market is pricing a commodity assumption the commodity market does not share.** At C$32.65, the valuation is consistent with WTI near US$85 held indefinitely. That price contains a geopolitical risk premium and the futures curve is in backwardation. At a US$70 mid-cycle price, NAV falls to approximately C$26.

**Paramount is simultaneously expensive and cheap.** Against Shell's April 2026 acquisition of ARC Resources, it trades at a 50% premium per flowing barrel and a 26% discount per 2P barrel. You are paying a premium for barrels that exist in the reserve report but are not yet flowing.

---

## Repository structure

```
├── Paramount_Resources_Initiating_Coverage.pdf   # The deliverable
├── Paramount_Resources_Initiating_Coverage.md    # Source for the note
│
├── models/
│   ├── POU_DataBook_v1.xlsx          # Module 1 — reconciled disclosure
│   ├── POU_OperatingModel_v1.xlsx    # Module 2 — 5-year forecast, 3 scenarios
│   ├── POU_NAV_v1.xlsx               # Module 4 — sum-of-the-parts valuation
│   └── POU_Comps_v1.xlsx             # Module 5 — trading & transaction comps
│
├── analysis/
│   ├── type_curve.py                 # Module 3 — decline curves & Monte Carlo
│   ├── wg_type_curve.png             # P90/P50/P10 decline and cumulative recovery
│   ├── wg_economics.png              # NPV, IRR and EUR distributions
│   ├── wg_tornado.png                # Sensitivity ranking of value drivers
│   └── wg_simulations.csv            # 8,000 simulation paths
│
└── build/                            # Scripts that generate the workbooks
    ├── build_databook.py
    ├── build_opmodel.py
    ├── build_nav.py
    └── build_comps.py
```

Every Excel workbook is **generated from code**, not built by hand. Rerunning a build script reproduces the workbook exactly, which makes the models auditable and version-controllable.

---

## The modules

### 1 — Data Book
Reconciles all reported financials, production, netbacks, reserves, guidance, hedges and capital structure into a single source of truth. Every calculated total is a formula and ties to Paramount's reported figures: netbacks to the cent across six periods, capital expenditure exactly, reserves to 0.01%.

### 2 — Operating Model
Five-year forecast (2026E–2030E) driven by asset-level production ramps and a switchable Bear/Base/Bull price deck. Includes automatic sanity checks against company guidance. The bear case shows net debt reaching C$730mm against C$750mm of undrawn facilities — the scenario in which the equity depends on capital markets access rather than operations.

### 3 — Type Curve & Well Economics *(Python)*
The analytical core. Arps hyperbolic decline

$$q(t) = \frac{q_i}{(1 + b D_i t)^{1/b}}$$

with a terminal exponential switchover (without which cumulative recovery diverges for *b* ≥ 1), calibrated to the company's disclosed 210-day rate. Uncertainty is propagated through **8,000 Monte Carlo paths** with correlated commodity draws, lognormal well-quality dispersion and a Beta-distributed uptime factor.

**Validation:** the well cost implied by the disclosed ~8-month payout is C$13.9mm against C$13.0mm assumed independently — a 7% gap between two unrelated routes.

| | P90 (low) | P50 | P10 (high) |
|---|---|---|---|
| EUR (MBoe) | 910 | 1,376 | 2,055 |
| NPV10 (C$mm) | 8.5 | 21.3 | 40.6 |
| Payout (months) | 5 | 10 | 22 |

### 4 — Net Asset Value
Sum-of-the-parts using closed-form annuity mathematics so the sensitivity grid stays live. Willesden Green is valued off the Module 3 type curve; Sinclair as a risked development project; Kaybob as a declining annuity; land at a per-acre mark. Corporate overhead is capitalised and long-dated abandonment deducted — both frequently omitted.

**NAV C$31.65 per share.** Willesden Green is 83% of gross value; Sinclair, despite dominating the narrative, is 17%.

### 5 — Comparables
Trading multiples against Canadian gas-weighted peers, plus control transactions. Two of the closest comparables were acquired during 2026 — ARC Resources by Shell (~C$22bn, April) and NuVista by Ovintiv (~C$3.8bn, February).

---

## Reproducing the analysis

```bash
pip install -r requirements.txt

python build/build_databook.py     # generates the Data Book
python build/build_opmodel.py      # generates the Operating Model
python build/build_nav.py          # generates the NAV model
python build/build_comps.py        # generates the Comparables model
python analysis/type_curve.py      # runs the Monte Carlo, writes charts + CSV
python build/make_pdf.py           # renders the research note to PDF
```

---

## Known limitations

Stated deliberately, because a model whose weaknesses are undocumented is harder to trust than one whose weaknesses are named.

- **The type curve is calibrated, not fitted.** Paramount does not publish well-level time series, so *b* and *Dᵢ* carry informative priors rather than being estimated from data. The appropriate upgrade is nonlinear least squares against monthly well data from the Alberta Energy Regulator or Petrinex.
- **Peer market data requires refreshing.** Share prices, share counts and net debt for the comparable set are flagged `VERIFY` in the workbook. Production, cash flow and transaction figures are sourced and reliable.
- **Modelled downside is too benign.** Parameter uncertainty is modelled; scenario uncertainty (plant outages, egress interruptions, geological surprises) is not. Disclosed well results are also a selection-biased sample.
- **The drilling programme assumes no well interference.** Parent-child depletion is an active issue in the Duvernay, so 16 wells per year is a floor.
- **Land value is a judgement.** Removing it entirely reduces NAV by C$2.42 per share.
- **Tax pools are approximated** by a single effective drag rather than scheduled explicitly.

---

## Sources

Paramount's Q2 2026 results (6 August 2026), FY2025 results (3 March 2026), the 2025 Annual Information Form including the McDaniel & Associates reserve report effective 31 December 2025, and filings on SEDAR+. Transaction data from acquirer and target announcements. Market data as at 21 August 2026.

---

## Tools

Python (NumPy, SciPy, pandas, Matplotlib, openpyxl) · Excel · pandoc + wkhtmltopdf
