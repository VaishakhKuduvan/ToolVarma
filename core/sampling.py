"""
Sampling engine — SA 530. Rule-based + risk-weighted selection over a Trial Balance
and a Day Book. Moved from the Integrated Audit Tool (app.py); the UI is pages/3_Sampling.py.
"""
import numpy as np
import pandas as pd

BASE_SIZES = {  # sample size per voucher category by risk level
    "High":   {"Journal": 25, "Cash": 20, "Credit/Debit Note": 15, "Bank": 20, "Sales": 25, "Purchase": 25, "Other": 15},
    "Medium": {"Journal": 15, "Cash": 12, "Credit/Debit Note": 10, "Bank": 12, "Sales": 15, "Purchase": 15, "Other": 10},
    "Low":    {"Journal": 10, "Cash": 8,  "Credit/Debit Note": 6,  "Bank": 8,  "Sales": 10, "Purchase": 10, "Other": 6},
}
TB_TOTAL = {"High": 40, "Medium": 30, "Low": 20}
TB_REQUIRED = ["SCHEDULE III MAIN GROUP", "SCHEDULE III SUB GROUP", "COMPONENT", "ledger_name", "tally_main_group",
               "tally_sub_group", "net_opening_balance", "Debit", "Credit", "net_closing_balance"]
DB_REQUIRED = ["Vch Type", "Debit Amount", "Credit Amount"]
VAGUE = ["adjustment", "misc", "to be", "provisional", "as per advice", "reversal", "correction"]
FLOOR = 50000


def weighted_sample(df, n, col):
    """Value-weighted (MUS-style) sample without replacement, reproducible (seed 42)."""
    if len(df) == 0 or n <= 0:
        return df.iloc[0:0]
    n = min(n, len(df))
    w = df[col].abs().astype(float).to_numpy()
    w = np.nan_to_num(w, nan=0.0)
    if n >= len(df):
        return df
    nonzero = int((w > 0).sum())
    if nonzero < n:  # not enough weighted items: take them all, top up randomly
        take = np.flatnonzero(w > 0)
        rest = np.setdiff1d(np.arange(len(df)), take)
        rng = np.random.RandomState(42)
        extra = rng.choice(rest, n - len(take), replace=False)
        return df.iloc[np.concatenate([take, extra])]
    idx = np.random.RandomState(42).choice(len(df), size=n, replace=False, p=w / w.sum())
    return df.iloc[idx]


def classify_vch(v):
    t = str(v).lower()
    for kw, name in [("journal", "Journal"), ("cash", "Cash"), ("credit note", "Credit/Debit Note"),
                     ("debit note", "Credit/Debit Note"), ("bank", "Bank"), ("sales", "Sales"),
                     ("purchase", "Purchase")]:
        if kw in t:
            return name
    # Tally payment/receipt/contra are bank/cash-like but kept as Other per spec
    return "Other"


def missing_columns(df, required):
    return [c for c in required if c not in df.columns]


def sample_trial_balance(tb, risk, pm):
    """Returns (selected_df, stats). pm is in ₹."""
    tb = tb.copy()
    for c in ["net_opening_balance", "Debit", "Credit", "net_closing_balance"]:
        tb[c] = pd.to_numeric(tb[c], errors="coerce").fillna(0)
    tb["SCHEDULE III MAIN GROUP"] = tb["SCHEDULE III MAIN GROUP"].fillna("Unclassified")
    tb["abs_closing"] = tb["net_closing_balance"].abs()
    tb["abs_activity"] = tb["Debit"].abs() + tb["Credit"].abs()
    tb["risk_indicator"] = tb[["abs_closing", "abs_activity"]].max(axis=1)
    key_items = tb[tb["risk_indicator"] >= pm].copy()
    key_items["Selection Basis"] = "Above PM — 100% check"
    pool = tb[(tb["risk_indicator"] >= FLOOR) & (tb["risk_indicator"] < pm)]
    groups = pool["SCHEDULE III MAIN GROUP"].unique()
    per = max(2, TB_TOTAL[risk] // max(len(groups), 1))
    parts = [weighted_sample(pool[pool["SCHEDULE III MAIN GROUP"] == g], per, "risk_indicator") for g in groups]
    sampled = (pd.concat(parts) if parts else pool.iloc[0:0]).copy()
    sampled["Selection Basis"] = "MUS sample (stratified)"
    sel = pd.concat([key_items, sampled]).drop_duplicates()
    stats = {"ledgers": len(tb), "key": len(key_items), "sampled": len(sampled), "selected": len(sel)}
    return sel, stats


def sample_daybook(db, risk, pm):
    """Returns (selected_df, summary_df). pm is in ₹."""
    db = db.copy()
    db["Debit Amount"] = pd.to_numeric(db["Debit Amount"], errors="coerce").fillna(0)
    db["Credit Amount"] = pd.to_numeric(db["Credit Amount"], errors="coerce").fillna(0)
    db["amount"] = db[["Debit Amount", "Credit Amount"]].abs().max(axis=1)
    db["Stratum"] = db["Vch Type"].apply(classify_vch)
    dedupe = [c for c in ["Vch No.", "Date"] if c in db.columns]
    out, summary = [], []
    for stratum, g in db.groupby("Stratum"):
        key = g[g["amount"] >= pm].copy()
        key["Selection Basis"] = "Above PM — 100% check"
        sub = g[g["amount"] < pm]
        flag_mask = pd.Series(False, index=sub.index)
        if stratum == "Journal":
            flag_mask |= (sub["amount"] % 100000 == 0) & (sub["amount"] >= 100000)
            if "Day" in sub.columns:
                flag_mask |= sub["Day"].astype(str).str.strip().str.lower().str[:3].isin(["sat", "sun"]) \
                             & (sub["amount"] >= FLOOR)
        if "Narrations" in sub.columns:
            pat = "|".join(VAGUE)
            flag_mask |= sub["Narrations"].astype(str).str.lower().str.contains(pat, na=False) \
                         & (sub["amount"] >= FLOOR)
        flags = sub[flag_mask].copy()
        flags["Selection Basis"] = "Red flag"
        pool = sub[(~flag_mask) & (sub["amount"] >= FLOOR)]
        n = min(BASE_SIZES[risk][stratum], len(pool))
        if stratum in ("Sales", "Purchase", "Bank"):
            smp = weighted_sample(pool, n, "amount")
        else:
            smp = pool.sample(n=n, random_state=42) if n > 0 else pool.iloc[0:0]
        smp = smp.copy()
        smp["Selection Basis"] = "Sample"
        comb = pd.concat([key, flags, smp])
        if dedupe:
            comb = comb.drop_duplicates(subset=dedupe)
        out.append(comb)
        summary.append({"Stratum": stratum, "Population": len(g), "Key Items": len(key),
                        "Red Flags": len(flags), "Sampled": len(smp), "Total Selected": len(comb)})
    sel = pd.concat(out) if out else db.iloc[0:0]
    return sel, pd.DataFrame(summary)
