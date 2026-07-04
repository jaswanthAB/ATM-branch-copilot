"""
Phase 1: Synthetic Branch & ATM Data Generator
Diebold Nixdorf Advisory Services — Case Study Project

Generates ~75 branches over 18 months with realistic patterns:
- Seasonal transaction variation
- ~10-12 branches with genuine performance problems
- ATM-specific metrics
- Natural noise and occasional missing values
"""

import numpy as np
import pandas as pd
from datetime import date
import random

np.random.seed(42)
random.seed(42)

# ── Configuration ─────────────────────────────────────────────────────────────
N_BRANCHES = 75
MONTHS = pd.date_range("2023-01-01", periods=18, freq="MS")

REGIONS = {
    "Northeast": 18,
    "Southeast": 15,
    "Midwest":   17,
    "Southwest": 13,
    "West":      12,
}

BRANCH_TYPES = ["Full-Service", "ATM-Only", "In-Store"]
BRANCH_TYPE_WEIGHTS = [0.55, 0.25, 0.20]

STATES_BY_REGION = {
    "Northeast": ["NY", "MA", "PA", "CT", "NJ"],
    "Southeast": ["FL", "GA", "NC", "VA", "SC"],
    "Midwest":   ["IL", "OH", "MI", "IN", "WI"],
    "Southwest": ["TX", "AZ", "NM", "NV", "CO"],
    "West":      ["CA", "WA", "OR", "UT", "ID"],
}

# Branches deliberately set up to have problems (indices chosen after generation)
UNDERPERFORMER_FRACTION = 0.15   # ~11 branches
ATM_PROBLEM_FRACTION    = 0.12   # ~9 branches (can overlap)

# ── Branch master ─────────────────────────────────────────────────────────────
def build_branch_master():
    branch_ids, regions, states, branch_types = [], [], [], []
    bid = 1001
    for region, count in REGIONS.items():
        for _ in range(count):
            branch_ids.append(f"BR-{bid}")
            regions.append(region)
            states.append(random.choice(STATES_BY_REGION[region]))
            branch_types.append(
                random.choices(BRANCH_TYPES, weights=BRANCH_TYPE_WEIGHTS, k=1)[0]
            )
            bid += 1

    df = pd.DataFrame({
        "Branch_ID":   branch_ids,
        "Region":      regions,
        "State":       states,
        "Branch_Type": branch_types,
    })

    # Base performance tier — hidden driver of variation
    df["_perf_tier"] = np.random.choice(
        ["high", "medium", "low"], size=len(df), p=[0.30, 0.55, 0.15]
    )

    # Flag underperformers and ATM-problem branches
    n_under = int(len(df) * UNDERPERFORMER_FRACTION)
    n_atm   = int(len(df) * ATM_PROBLEM_FRACTION)
    low_idx  = df.index[df["_perf_tier"] == "low"].tolist()

    # Ensure we have enough — supplement with random medium-tier if needed
    under_idx = random.sample(low_idx, min(n_under, len(low_idx)))
    remaining = [i for i in df.index if i not in under_idx]
    if len(under_idx) < n_under:
        under_idx += random.sample(remaining, n_under - len(under_idx))

    atm_pool  = [i for i in df.index if df.loc[i, "Branch_Type"] != "Full-Service" or True]
    atm_idx   = random.sample(atm_pool, n_atm)

    df["_underperformer"] = df.index.isin(under_idx)
    df["_atm_problem"]    = df.index.isin(atm_idx)

    return df


# ── Monthly seasonality curve ─────────────────────────────────────────────────
def seasonal_factor(month: int) -> float:
    """Returns a multiplier 0.85–1.15 based on month (1-indexed)."""
    # Jan low, summer moderate, Dec peak (retail banking / holiday pattern)
    base = [0.87, 0.88, 0.93, 0.97, 1.00, 1.03,
            1.05, 1.04, 1.00, 0.98, 0.97, 1.15]
    return base[month - 1]


# ── Monthly records ───────────────────────────────────────────────────────────
def build_monthly_records(master: pd.DataFrame) -> pd.DataFrame:
    records = []

    for _, branch in master.iterrows():
        is_under   = branch["_underperformer"]
        is_atm_bad = branch["_atm_problem"]
        btype      = branch["Branch_Type"]
        tier       = branch["_perf_tier"]

        # Base transaction volume (monthly)
        base_txn = {"high": 14_000, "medium": 9_000, "low": 4_500}[tier]
        if btype == "ATM-Only":
            base_txn *= 0.45
        elif btype == "In-Store":
            base_txn *= 0.70

        # Declining trend for underperformers (compound monthly decline)
        decline_rate = np.random.uniform(0.015, 0.030) if is_under else 0.0

        # Base ATM uptime
        base_uptime = np.random.uniform(0.88, 0.97) if is_atm_bad else np.random.uniform(0.96, 0.999)

        # Base satisfaction (0-100)
        base_sat = {"high": 82, "medium": 74, "low": 63}[tier]

        # Operating costs (monthly, USD)
        base_cost = {"Full-Service": 95_000, "ATM-Only": 12_000, "In-Store": 35_000}[btype]
        base_cost *= np.random.uniform(0.85, 1.15)  # branch-specific cost variation

        for t, month_dt in enumerate(MONTHS):
            season = seasonal_factor(month_dt.month)

            # ── Transactions ──────────────────────────────────────────────────
            trend_factor = (1 - decline_rate) ** t if is_under else 1.0
            txn = int(base_txn * season * trend_factor * np.random.uniform(0.93, 1.07))

            # ── Foot traffic (correlated with txn) ───────────────────────────
            if btype == "ATM-Only":
                foot = None  # no meaningful foot traffic for ATM-only
            else:
                foot = int(txn * np.random.uniform(0.08, 0.14))

            # ── Wait time (inverse relationship with txn, noisier for struggling) ──
            if btype == "ATM-Only":
                wait = None
            else:
                base_wait = 4.2 if tier == "high" else (6.1 if tier == "medium" else 9.0)
                wait = round(base_wait + np.random.normal(0, 0.8), 1)
                wait = max(1.0, wait)

            # ── ATM metrics ───────────────────────────────────────────────────
            if btype == "Full-Service":
                # Full-service branches have 1-3 ATMs
                n_atms = random.randint(1, 3)
            elif btype == "ATM-Only":
                n_atms = random.randint(2, 5)
            else:
                n_atms = 1

            # Uptime degrades over time for atm_problem branches
            atm_uptime_drift = -0.005 * t if is_atm_bad else 0.0
            atm_uptime = round(
                min(1.0, max(0.50, base_uptime + atm_uptime_drift + np.random.normal(0, 0.012))),
                4
            )
            downtime_incidents = 0 if atm_uptime > 0.97 else random.randint(1, 5) if atm_uptime < 0.90 else random.randint(0, 2)
            atm_txn_per_day = round(txn / 30 / n_atms * np.random.uniform(0.9, 1.1), 1)

            # Cash level (% of capacity, refilled roughly weekly)
            cash_level_pct = round(np.random.uniform(0.25, 0.95), 2)

            # ── Customer satisfaction ─────────────────────────────────────────
            sat_drift = -0.4 * t if is_under else 0.0
            csat = round(
                min(100, max(20, base_sat + sat_drift + np.random.normal(0, 3.5))),
                1
            )

            # ── Revenue & costs ───────────────────────────────────────────────
            # Bank branch revenue = transaction fees + account fees + loan/deposit income
            # Full-Service: $120k-$220k/month; In-Store: $40k-$80k; ATM-Only: $8k-$20k
            base_revenue = {
                "Full-Service": np.random.uniform(120_000, 220_000),
                "In-Store":     np.random.uniform(40_000,   80_000),
                "ATM-Only":     np.random.uniform(8_000,    20_000),
            }[btype]
            # Underperformers earn less; revenue also tracks with transaction trend
            revenue_multiplier = trend_factor * season * np.random.uniform(0.92, 1.08)
            if is_under:
                revenue_multiplier *= np.random.uniform(0.70, 0.88)
            revenue = round(base_revenue * revenue_multiplier, 2)
            cost    = round(base_cost * season * np.random.uniform(0.97, 1.03), 2)
            roi_pct = round((revenue - cost) / cost * 100, 2)

            # ── Introduce ~5% missing values on select columns ────────────────
            def maybe_null(val, p=0.04):
                return None if np.random.random() < p else val

            records.append({
                "Branch_ID":             branch["Branch_ID"],
                "Region":                branch["Region"],
                "State":                 branch["State"],
                "Branch_Type":           branch["Branch_Type"],
                "Month":                 month_dt.strftime("%Y-%m-%d"),
                "Transaction_Volume":    txn,
                "Foot_Traffic":          maybe_null(foot),
                "Avg_Wait_Time_Min":     maybe_null(wait),
                "Num_ATMs":              n_atms,
                "ATM_Uptime_Pct":        atm_uptime,
                "ATM_Downtime_Incidents":downtime_incidents,
                "ATM_Txn_Per_Day":       maybe_null(atm_txn_per_day),
                "Cash_Level_Pct":        cash_level_pct,
                "CSAT_Score":            maybe_null(csat),
                "Monthly_Revenue_USD":   revenue,
                "Monthly_Cost_USD":      cost,
                "ROI_Pct":               roi_pct,
                # Hidden flags — used by pipeline but NOT included in final output
                "_is_underperformer":    is_under,
                "_is_atm_problem":       is_atm_bad,
            })

    return pd.DataFrame(records)


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Building branch master...")
    master = build_branch_master()
    print(f"  {len(master)} branches across {len(REGIONS)} regions")

    print("Generating monthly records...")
    df = build_monthly_records(master)
    print(f"  {len(df):,} total monthly records ({len(MONTHS)} months × {N_BRANCHES} branches)")

    # Drop hidden flags before saving raw file
    raw = df.drop(columns=["_is_underperformer", "_is_atm_problem"])

    out_path = "raw_branch_atm_data.csv"
    raw.to_csv(out_path, index=False)
    print(f"  Saved → {out_path}")

    # Quick sanity checks
    print("\nSanity checks:")
    print(f"  Branches: {raw['Branch_ID'].nunique()}")
    print(f"  Date range: {raw['Month'].min()} → {raw['Month'].max()}")
    print(f"  Avg Transaction Volume: {raw['Transaction_Volume'].mean():,.0f}")
    print(f"  Avg ATM Uptime: {raw['ATM_Uptime_Pct'].mean():.1%}")
    print(f"  Missing values:\n{raw.isnull().sum()[raw.isnull().sum()>0]}")

    # Save master with flags for pipeline use
    master.to_csv("branch_master_with_flags.csv", index=False)
    print("\nDone. Files written: raw_branch_atm_data.csv, branch_master_with_flags.csv")
