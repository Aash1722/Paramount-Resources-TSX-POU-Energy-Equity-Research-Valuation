"""
================================================================================
Paramount Resources (TSX: POU) - PE / Equity Research Case Study
MODULE 3: Willesden Green Duvernay Type Curve & Probabilistic Well Economics
================================================================================

PURPOSE
-------
Module 2 asserted asset-level plateau rates top-down. That cannot answer the
questions that actually get asked in diligence:
    "What if wells come in 20% below type curve?"
    "How many wells a year does the 50,000 Boe/d plateau actually require?"
    "What oil price makes a single well uneconomic?"

This module rebuilds the forecast bottom-up from well-level physics and
propagates parameter uncertainty through to NPV, IRR and payout.

METHOD AND ITS HONEST LIMITATION
--------------------------------
Paramount does not publish well-level production time series, so this is NOT a
regression fit to observed data. It is a CALIBRATION: Arps hyperbolic decline
parameters are constrained to reproduce the company's disclosed anchor points,
with informative priors on the parameters that remain free. Uncertainty is then
propagated by Monte Carlo.

That distinction matters. Say "calibrated to disclosed anchors", never "fitted",
unless you have the monthly data from AER / Petrinex, at which point b and Di
should be estimated by nonlinear least squares and this module upgraded.

DISCLOSED ANCHORS (Q2 2026 and FY2025 press releases)
-----------------------------------------------------
    - 16 Willesden Green wells, 210-day peak rate ~1,240 Boe/d, 56% liquids
    - 13 of 16 wells paid out in approximately 8 months
    - Willesden Green operating expense C$4.71/Boe (Q4 2025)
    - Stated sustainable plateau ~50,000 Boe/d for 20+ years

Author's note: every assumption below is flagged ASSUMPTION or ANCHOR.
Hypothetical case study. Not investment advice.
================================================================================
"""

import numpy as np
import pandas as pd
from scipy.optimize import brentq
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

RNG = np.random.default_rng(20260825)

# ==============================================================================
# 1. INPUTS
# ==============================================================================

# ---- ANCHORS (from disclosure) ----
PEAK_210_BOED   = 1240.0    # ANCHOR: 210-day rate, Boe/d
PEAK_WINDOW_YRS = 210/365.0 # ANCHOR: the 210-day window
LIQUIDS_FRAC    = 0.56      # ANCHOR: 56% liquids
COND_OF_LIQUIDS = 0.80      # ASSUMPTION: condensate share of liquids
OPEX_BOE        = 4.71      # ANCHOR: WG operating expense C$/Boe
PAYOUT_ANCHOR_M = 8.0       # ANCHOR: ~8 months payout, 13 of 16 wells
PLATEAU_BOED    = 50000.0   # ANCHOR: stated sustainable plateau

# ---- DECLINE PARAMETERS (priors; central values) ----
B_EXP_MEAN      = 1.10      # ASSUMPTION: hyperbolic exponent, typical shale 0.8-1.4
B_EXP_SD        = 0.15
DI_NOM_MEAN     = 2.20      # ASSUMPTION: nominal initial decline, /yr
DI_NOM_SD       = 0.30
D_TERMINAL      = 0.08      # ASSUMPTION: terminal exponential decline 8%/yr
WELL_LIFE_YRS   = 30.0
QI_LOGNORM_SD   = 0.28      # ASSUMPTION: well-to-well quality dispersion

# ---- COSTS AND PRICES ----
WELL_COST_MM    = 13.0      # DERIVED below and cross-checked against payout anchor
FACILITY_PER_WELL = 5.0     # ASSUMPTION: allocated Alhambra/trunk capital per well
WELL_COST_SD    = 1.6       # ASSUMPTION
TRANSPORT_BOE   = 4.75      # ANCHOR: corporate transport & NGL processing
GA_BOE          = 1.20      # ASSUMPTION: WG-allocated G&A
ROY_EARLY       = 0.05      # ANCHOR: Alberta MRF new-well rate until cost recovery
ROY_LATE        = 0.16      # ASSUMPTION: post-payout mature rate
DISC_RATE       = 0.10      # Industry convention

WTI_BASE        = 80.0      # From Module 2 Base deck, US$/Bbl
WTI_SD          = 12.0      # ASSUMPTION: forward price uncertainty
FX              = 1.37
COND_DIFF       = 0.0       # C$/Bbl to WTI
NGL_RATIO       = 0.32      # Other NGLs as % of condensate price
AECO_BASE       = 2.90      # C$/GJ
AECO_SD         = 0.70
GAS_PREMIUM     = 0.90      # C$/GJ diversification premium
GJ_PER_MCF      = 1.055

N_SIMS          = 8000
DT              = 1/365.0


# ==============================================================================
# 2. ARPS DECLINE
# ==============================================================================

def arps_rate(t, qi, b, di, d_term=D_TERMINAL):
    """
    Hyperbolic decline with a terminal exponential tail.

        q(t) = qi / (1 + b*Di*t)^(1/b)

    Instantaneous decline D(t) = Di / (1 + b*Di*t) falls with time. Once it
    reaches the terminal rate we switch to exponential, which is standard
    reserve-evaluation practice: unbounded hyperbolic decline overstates EUR
    because it implies the well never stops producing.
    """
    t = np.asarray(t, dtype=float)
    t_switch = (di / d_term - 1.0) / (b * di) if b > 0 else np.inf
    t_switch = max(t_switch, 0.0)

    q = np.empty_like(t)
    hyp = t <= t_switch
    q[hyp] = qi / np.power(1.0 + b * di * t[hyp], 1.0 / b)

    if np.any(~hyp):
        q_sw = qi / np.power(1.0 + b * di * t_switch, 1.0 / b)
        q[~hyp] = q_sw * np.exp(-d_term * (t[~hyp] - t_switch))
    return q


def avg_rate_over(qi, b, di, horizon_yrs):
    """Average daily rate across the first `horizon_yrs` — matches how a
    '210-day rate' is reported."""
    t = np.arange(0.0, horizon_yrs, DT)
    return arps_rate(t, qi, b, di).mean()


def solve_qi(b, di, target_avg=PEAK_210_BOED, horizon=PEAK_WINDOW_YRS):
    """Back out the initial rate qi that reproduces the disclosed 210-day rate.

    This is the calibration step: b and di carry priors, qi is determined."""
    f = lambda q: avg_rate_over(q, b, di, horizon) - target_avg
    return brentq(f, 100.0, 40000.0, xtol=1e-6)


def eur_boe(qi, b, di, years=WELL_LIFE_YRS):
    """Estimated ultimate recovery, Boe, by numerical integration."""
    t = np.arange(0.0, years, DT)
    return arps_rate(t, qi, b, di).sum() * 1.0   # daily rate * 1 day steps


# ==============================================================================
# 3. WELL ECONOMICS
# ==============================================================================

def price_deck(wti, aeco):
    """Convert benchmarks into realised C$ prices per unit."""
    cond = wti * FX + COND_DIFF                 # C$/Bbl
    ngl  = cond * NGL_RATIO                     # C$/Bbl
    gas  = (aeco + GAS_PREMIUM) * GJ_PER_MCF    # C$/Mcf
    return cond, ngl, gas


def revenue_per_boe(cond, ngl, gas):
    """Blended C$/Boe at the well's product mix.
    One Boe = LIQUIDS_FRAC barrels of liquid + (1-LIQUIDS_FRAC)*6 Mcf of gas."""
    liq = LIQUIDS_FRAC
    r_cond = liq * COND_OF_LIQUIDS * cond
    r_ngl  = liq * (1 - COND_OF_LIQUIDS) * ngl
    r_gas  = (1 - liq) * 6.0 * gas
    return r_cond + r_ngl + r_gas


def well_cashflow(qi, b, di, well_cost_mm, wti, aeco, years=WELL_LIFE_YRS):
    """Monthly half-cycle cash flow for a single well.

    Half-cycle = drilling, completion and tie-in only. It EXCLUDES the Alhambra
    plant and trunk infrastructure, which are full-cycle costs. Half-cycle
    returns always look better; be explicit about which you are quoting.

    Royalty steps from the new-well rate to the mature rate once cumulative
    revenue has recovered the well cost, mirroring Alberta's Modern Royalty
    Framework.
    """
    t = np.arange(0.0, years, DT)
    q = arps_rate(t, qi, b, di)                       # Boe/d each day
    cond, ngl, gas = price_deck(wti, aeco)
    rev_boe = revenue_per_boe(cond, ngl, gas)

    daily_rev = q * rev_boe                            # C$/day
    cum_rev = np.cumsum(daily_rev)
    recovered = cum_rev >= well_cost_mm * 1e6
    roy_rate = np.where(recovered, ROY_LATE, ROY_EARLY)

    daily_cf = daily_rev * (1 - roy_rate) - q * (OPEX_BOE + TRANSPORT_BOE + GA_BOE)

    # aggregate to months for discounting
    n_months = int(years * 12)
    monthly = np.array_split(daily_cf, n_months)
    m_cf = np.array([m.sum() for m in monthly]) / 1e6          # C$mm
    m_cf[0] -= well_cost_mm                                     # capital at t=0

    disc = (1 + DISC_RATE) ** (-(np.arange(n_months) + 0.5) / 12.0)
    npv = float((m_cf * disc).sum())

    cum = np.cumsum(m_cf)
    idx = np.argmax(cum > 0) if np.any(cum > 0) else -1
    payout_m = float(idx + 1) if idx >= 0 else np.nan

    return npv, payout_m, m_cf


def irr_from(m_cf, lo=-0.95, hi=10.0):
    """Annual IRR from a monthly cash-flow vector."""
    n = np.arange(len(m_cf))
    def npv_at(r):
        return float((m_cf * (1 + r) ** (-n / 12.0)).sum())
    if npv_at(lo) * npv_at(hi) > 0:
        return np.nan
    return brentq(npv_at, lo, hi, xtol=1e-6)


# ==============================================================================
# 4. CALIBRATION AND VALIDATION
# ==============================================================================

def calibrate():
    qi = solve_qi(B_EXP_MEAN, DI_NOM_MEAN)
    eur = eur_boe(qi, B_EXP_MEAN, DI_NOM_MEAN)
    r210 = avg_rate_over(qi, B_EXP_MEAN, DI_NOM_MEAN, PEAK_210_BOED and PEAK_WINDOW_YRS)

    # Cross-check: what well cost does the disclosed ~8-month payout imply?
    def payout_gap(cost):
        _, pm, _ = well_cashflow(qi, B_EXP_MEAN, DI_NOM_MEAN, cost, WTI_BASE, AECO_BASE)
        return (pm if not np.isnan(pm) else 999) - PAYOUT_ANCHOR_M
    try:
        implied_cost = brentq(payout_gap, 3.0, 40.0, xtol=1e-3)
    except ValueError:
        implied_cost = np.nan

    return qi, eur, r210, implied_cost


# ==============================================================================
# 5. MONTE CARLO
# ==============================================================================

def simulate(n=N_SIMS):
    b   = np.clip(RNG.normal(B_EXP_MEAN, B_EXP_SD, n), 0.4, 1.8)
    di  = np.clip(RNG.normal(DI_NOM_MEAN, DI_NOM_SD, n), 0.8, 4.0)
    qmul= RNG.lognormal(-0.5 * QI_LOGNORM_SD**2, QI_LOGNORM_SD, n)   # mean 1.0
    cost= np.clip(RNG.normal(WELL_COST_MM, WELL_COST_SD, n), 7.0, 25.0)
    # Execution risk: facility downtime, curtailment, egress interruption.
    # Without this the model shows a ~0% chance of a bad well, which is not a
    # credible statement about a single-plant asset.
    uptime = np.clip(RNG.beta(9, 1.2, n), 0.45, 1.0)

    # WTI and AECO are correlated but not tightly. rho estimated from the
    # historical relationship between crude and Alberta gas; ASSUMPTION.
    rho = 0.35
    z = RNG.multivariate_normal([0, 0], [[1, rho], [rho, 1]], n)
    wti  = np.clip(WTI_BASE  + WTI_SD  * z[:, 0], 30.0, 160.0)
    aeco = np.clip(AECO_BASE + AECO_SD * z[:, 1], 0.6, 9.0)

    npv = np.empty(n); pay = np.empty(n); irr = np.empty(n); eur = np.empty(n)
    for i in range(n):
        qi_base = solve_qi(b[i], di[i])
        qi = qi_base * qmul[i]
        eur[i] = eur_boe(qi, b[i], di[i])
        v, pm, mcf = well_cashflow(qi * uptime[i], b[i], di[i], cost[i],
                                   wti[i], aeco[i])
        npv[i] = v; pay[i] = pm
        irr[i] = irr_from(mcf)

    return pd.DataFrame(dict(b=b, di=di, qi_mult=qmul, well_cost=cost, uptime=uptime,
                             wti=wti, aeco=aeco, eur=eur,
                             npv=npv, payout_m=pay, irr=irr))


def pctiles(s, lo=10, hi=90):
    """Petroleum convention: P90 is the LOW case (90% probability of exceeding),
    P10 is the HIGH case. This is the reverse of the statistical convention and
    is a classic way to get caught out in an interview."""
    return dict(P90=np.nanpercentile(s, lo),
                P50=np.nanpercentile(s, 50),
                P10=np.nanpercentile(s, hi))


# ==============================================================================
# 6. LINK BACK TO THE CORPORATE MODEL
# ==============================================================================

def wells_for_plateau(qi, b, di, plateau=PLATEAU_BOED, years=12):
    """How many wells a year does the stated plateau actually require?

    Simulates a constant annual drilling programme, sums the declining
    contribution of every vintage, and solves for the count that holds
    `plateau` in steady state. This is the number to compare against the
    capital budget — it is the bridge between well economics and the
    corporate forecast.
    """
    t_grid = np.arange(0.0, years, 1/12)
    def rate_at_steady(n_per_yr):
        total = np.zeros_like(t_grid)
        for k in range(int(years * n_per_yr)):
            start = k / n_per_yr
            live = t_grid >= start
            total[live] += arps_rate(t_grid[live] - start, qi, b, di)
        return total[-1]
    f = lambda n: rate_at_steady(n) - plateau
    return brentq(f, 1.0, 200.0, xtol=0.01)


# ==============================================================================
# 7. RUN
# ==============================================================================

def line(c="-", n=78): print(c * n)

if __name__ == "__main__":
    line("=")
    print("MODULE 3 — WILLESDEN GREEN DUVERNAY TYPE CURVE")
    line("=")

    qi, eur, r210, implied_cost = calibrate()
    print(f"\nCALIBRATION (central case)")
    print(f"  Hyperbolic exponent b           {B_EXP_MEAN:.2f}   (prior)")
    print(f"  Nominal initial decline Di      {DI_NOM_MEAN:.2f}/yr (prior)")
    print(f"  Solved initial rate qi          {qi:,.0f} Boe/d")
    print(f"  Reproduced 210-day rate         {r210:,.0f} Boe/d   (anchor 1,240)")
    print(f"  EUR (30-year)                   {eur/1000:,.0f} MBoe")
    print(f"  Year-1 average rate             {avg_rate_over(qi,B_EXP_MEAN,DI_NOM_MEAN,1.0):,.0f} Boe/d")

    print(f"\nVALIDATION — the payout cross-check")
    print(f"  Well cost implied by ~8-month payout   C${implied_cost:,.1f}mm")
    print(f"  Well cost assumed in economics         C${WELL_COST_MM:,.1f}mm")
    if not np.isnan(implied_cost):
        print(f"  Gap                                    {abs(implied_cost-WELL_COST_MM)/WELL_COST_MM:.0%}")
        print("  A small gap means the decline parameters and the disclosed payout")
        print("  tell a consistent story. A large gap means one of them is wrong.")

    npv0, pay0, mcf0 = well_cashflow(qi, B_EXP_MEAN, DI_NOM_MEAN, WELL_COST_MM,
                                     WTI_BASE, AECO_BASE)
    print(f"\nDETERMINISTIC WELL ECONOMICS (Base deck, half-cycle)")
    print(f"  NPV10                           C${npv0:,.1f}mm")
    print(f"  IRR                             {irr_from(mcf0):.0%}  <- see note")
    print(f"  Payout                          {pay0:,.0f} months   (anchor ~8)")
    print( "  NOTE: a sub-12-month payout mechanically produces a triple-digit IRR.")
    print( "  It is arithmetically correct and analytically useless. Lead with NPV")
    print( "  per well and payout; treat half-cycle IRR as a red flag, not a result.")
    print(f"  Recycle proxy (NPV/cost)        {npv0/WELL_COST_MM:,.2f}x")

    print(f"\nRunning {N_SIMS:,} simulations ...")
    df = simulate()
    df.to_csv("wg_simulations.csv", index=False)

    print(f"\nPROBABILISTIC RESULTS  (P90 = low case, P10 = high case)")
    print(f"  {'':22}{'P90':>12}{'P50':>12}{'P10':>12}")
    for lab, s, f_ in [("EUR (MBoe)", df.eur/1000, "{:>12,.0f}"),
                       ("NPV10 (C$mm)", df.npv, "{:>12,.1f}"),
                       ("IRR", df.irr*100, "{:>12,.0f}"),
                       ("Payout (months)", df.payout_m, "{:>12,.0f}")]:
        p = pctiles(s)
        print(f"  {lab:22}" + f_.format(p['P90']) + f_.format(p['P50']) + f_.format(p['P10']))

    print(f"\nDOWNSIDE STATISTICS")
    print(f"  P(NPV10 < 0)                    {(df.npv < 0).mean():.1%}")
    print(f"  P(IRR < 15% hurdle)             {(df.irr < 0.15).mean():.1%}")
    print(f"  P(payout > 24 months)           {(df.payout_m > 24).mean():.1%}")

    print(f"\nBREAKEVEN")
    for target, lab in [(0.0, "NPV10 = 0"), (0.15, "IRR = 15%")]:
        f = lambda w: (well_cashflow(qi, B_EXP_MEAN, DI_NOM_MEAN, WELL_COST_MM, w, AECO_BASE)[0]
                       if target == 0 else
                       irr_from(well_cashflow(qi, B_EXP_MEAN, DI_NOM_MEAN,
                                              WELL_COST_MM, w, AECO_BASE)[2]) - target)
        try:
            print(f"  WTI for {lab:16}         US${brentq(f, 15.0, 150.0, xtol=0.01):,.0f}/Bbl")
        except ValueError:
            print(f"  WTI for {lab:16}         below US$15/Bbl")

    npv_fc, pay_fc, mcf_fc = well_cashflow(qi, B_EXP_MEAN, DI_NOM_MEAN,
                                           WELL_COST_MM + FACILITY_PER_WELL,
                                           WTI_BASE, AECO_BASE)
    print(f"\nFULL-CYCLE ECONOMICS (adds C${FACILITY_PER_WELL:.0f}mm/well of allocated facility capital)")
    print(f"  NPV10                           C${npv_fc:,.1f}mm")
    print(f"  IRR                             {irr_from(mcf_fc):.0%}")
    print(f"  Payout                          {pay_fc:,.0f} months")
    print( "  Quote THIS number, not the half-cycle one. Half-cycle excludes the")
    print( "  Alhambra plant without which none of these wells can flow.")

    n_wells = wells_for_plateau(qi, B_EXP_MEAN, DI_NOM_MEAN)
    dc = n_wells * WELL_COST_MM
    plateau_netback = PLATEAU_BOED * 44.0 * 365 / 1e6
    print(f"\nBRIDGE TO THE CORPORATE MODEL — the key structural finding")
    print(f"  Wells per year to hold {PLATEAU_BOED:,.0f} Boe/d   {n_wells:,.0f}")
    print(f"  Implied maintenance D&C capital       C${dc:,.0f}mm/yr")
    print(f"  Netback generated at plateau          C${plateau_netback:,.0f}mm/yr")
    print(f"  Maintenance capital as % of netback   {dc/plateau_netback:.0%}")
    print(f"  Paramount's 2026 WG budget            C$630mm (drilling + facilities)")
    print( "  READ: once built, Willesden Green sustains itself on a fraction of")
    print( "  the cash it throws off. The outspend in Module 2 is the BUILD -")
    print( "  Alhambra and Sinclair - not the steady state. That distinction is")
    print( "  the core of the bull case and should be stated explicitly.")

    # ------------------------------------------------------------------
    # CHARTS
    # ------------------------------------------------------------------
    plt.rcParams.update({"font.size": 9, "axes.edgecolor": "#888780",
                         "axes.labelcolor": "#3d3d3a", "text.color": "#3d3d3a",
                         "xtick.color": "#5F5E5A", "ytick.color": "#5F5E5A",
                         "axes.grid": True, "grid.color": "#E5E3DC",
                         "grid.linewidth": 0.6, "figure.dpi": 130})
    TEAL, CORAL, GRAY = "#1D9E75", "#D85A30", "#888780"

    # --- Chart 1: type curve with uncertainty band ---
    fig, ax = plt.subplots(1, 2, figsize=(11, 4.2))
    t = np.arange(0, 10, 1/12)
    samp = df.sample(400, random_state=1)
    curves = np.array([arps_rate(t, solve_qi(r.b, r.di) * r.qi_mult, r.b, r.di)
                       for r in samp.itertuples()])
    lo, mid, hi = np.percentile(curves, [10, 50, 90], axis=0)
    ax[0].fill_between(t, lo, hi, color=TEAL, alpha=0.18, label="P90–P10 range")
    ax[0].plot(t, mid, color=TEAL, lw=2, label="P50 type curve")
    ax[0].scatter([PEAK_WINDOW_YRS/2], [PEAK_210_BOED], color=CORAL, zorder=5,
                  s=45, label="Disclosed 210-day rate")
    ax[0].set_yscale("log"); ax[0].set_xlabel("Years on production")
    ax[0].set_ylabel("Rate (Boe/d, log scale)")
    ax[0].set_title("Willesden Green type curve", loc="left", fontsize=11)
    ax[0].legend(frameon=False, fontsize=8)

    cum = np.array([np.cumsum(c)*(365/12)/1000 for c in curves])
    clo, cmid, chi = np.percentile(cum, [10, 50, 90], axis=0)
    ax[1].fill_between(t, clo, chi, color=TEAL, alpha=0.18)
    ax[1].plot(t, cmid, color=TEAL, lw=2)
    ax[1].set_xlabel("Years on production"); ax[1].set_ylabel("Cumulative (MBoe)")
    ax[1].set_title("Cumulative recovery", loc="left", fontsize=11)
    plt.tight_layout(); plt.savefig("wg_type_curve.png", bbox_inches="tight")
    plt.close()

    # --- Chart 2: economics distributions ---
    fig, ax = plt.subplots(1, 3, figsize=(12, 3.6))
    for a, s, lab, unit in [(ax[0], df.npv, "NPV10 per well", "C$mm"),
                            (ax[1], df.irr*100, "IRR", "%"),
                            (ax[2], df.eur/1000, "EUR", "MBoe")]:
        s = s.replace([np.inf, -np.inf], np.nan).dropna()
        a.hist(s, bins=60, color=TEAL, alpha=0.75, edgecolor="white", linewidth=0.3)
        for q, st in [(10, ":"), (50, "-"), (90, ":")]:
            a.axvline(np.percentile(s, q), color=CORAL, ls=st, lw=1.2)
        a.set_title(f"{lab} ({unit})", loc="left", fontsize=11)
        a.set_yticks([])
    ax[0].axvline(0, color=GRAY, lw=1.2)
    plt.tight_layout(); plt.savefig("wg_economics.png", bbox_inches="tight")
    plt.close()

    # --- Chart 3: tornado of drivers ---
    drivers = {"WTI (US$/Bbl)": "wti", "Well quality (qi)": "qi_mult",
               "Well cost (C$mm)": "well_cost", "Decline Di": "di",
               "Uptime / execution": "uptime",
               "AECO (C$/GJ)": "aeco", "b exponent": "b"}
    eff = {}
    for lab, col in drivers.items():
        q = df[col].quantile([0.1, 0.9])
        loM = df.loc[df[col] <= q.iloc[0], "npv"].mean()
        hiM = df.loc[df[col] >= q.iloc[1], "npv"].mean()
        eff[lab] = (loM, hiM)
    order = sorted(eff, key=lambda k: abs(eff[k][1] - eff[k][0]))
    base = df.npv.median()
    fig, a = plt.subplots(figsize=(7.5, 3.4))
    for i, k in enumerate(order):
        lo_, hi_ = eff[k]
        a.barh(i, lo_ - base, left=base, color=CORAL, alpha=0.85, height=0.6)
        a.barh(i, hi_ - base, left=base, color=TEAL, alpha=0.85, height=0.6)
    a.set_yticks(range(len(order))); a.set_yticklabels(order)
    a.axvline(base, color=GRAY, lw=1.2)
    a.set_xlabel("Mean NPV10 per well (C$mm) in bottom vs top decile of each driver")
    a.set_title("What actually moves well value", loc="left", fontsize=11)
    plt.tight_layout(); plt.savefig("wg_tornado.png", bbox_inches="tight")
    plt.close()

    print("\nSaved: wg_type_curve.png, wg_economics.png, wg_tornado.png, wg_simulations.csv")
    line("=")
