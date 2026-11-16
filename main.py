#!/usr/bin/env python3
"""
Sector neighbour geometry + per-relation overshoot analysis.  (v2)

Matched to the real exports:
  source sheet : Code / SEC / latitude / longitude / Azimuth / Band   (no cell name)
  sector DB    : ... Site name / Site code / cellName / LAT / LONG / AZ / band / sec
  per-cell KPI : LNCEL name / LNBTS name + Avg UE distance, PRB usage per TTI DL,
                 Average CQI, Avg MCS PUSCH/PDSCH trans, PDCP SDU Volume,
                 Spectral efficiency - all combined in ONE export now
  neighbour HO : Source LNBTS name / Source LNCEL name / Target LNBTS name /
                 Target LNCEL name / Number of Inter eNB HO attempts per relation
                 (still its own independent export)

Stage 0   resolve  : source (Code, SEC, Band) -> real LNCELs in the DB
Stage 1   geometry : mutual-facing neighbours, same band only
Stage 2   cell KPI : UE distance, PRB, CQI, MCS PUSCH/PDSCH, traffic, spectral
                     efficiency - all per-cell KPIs from one combined export
                     (source + neighbour)
Stage 3   rel KPI  : inter-eNB HO attempts per neighbour relation (own file)
Stage 4   metrics  : overshoot index, fair-share ratio, ISD ratio, flags
Stage 5   export   : coloured Excel (relations + per-cell summary with cell-level flag)

Cell_Summary carries source-side PRB/CQI/MCS/traffic plus a worst-case
cell_severity and cell_overshoot_flag for one-click filtering.
"""

import os
import re

import numpy as np
import pandas as pd


# ============================================================
# CONFIG
# ============================================================

CONFIG = {
    # --- inputs -------------------------------------------------
    "source_csv": "data/source.csv",
    "sector_csv": "data/sector_db.csv",

    # combined per-cell KPI export: UE distance, PRB, CQI, MCS PUSCH/PDSCH,
    # traffic and spectral efficiency all live in this one sheet now.
    # Set to None to skip per-cell KPIs entirely.
    "percell_kpi_csv": "data/percell_kpi.csv",

    # inter-eNB HO attempts per neighbour relation - still its own
    # independent export, unaffected by the per-cell KPI merge.
    "rel_kpi_csv": "data/neighbor_ho.csv",

    # delimiters: None = auto-sniff among ; \t , |
    "source_sep": None,
    "sector_sep": None,
    "percell_kpi_sep": ";",
    "rel_sep":  ";",

    # set to "," if the board exports use decimal commas
    "decimal": ".",

    # --- output -------------------------------------------------
    "out_xlsx": "output/overshoot_analysis.xlsx",

    # --- geometry -----------------------------------------------
    "top_n": 5,
    "source_beamwidth": 50,
    "nbr_beamwidth": 50,
    "max_distance_km": 4,
    "distance_weight": 0.6,
    "azimuth_weight": 0.4,
    "same_band_only": True,
    "exclude_cosite": True,

    # "db" = take LAT/LONG/AZ from the sector DB for both ends (recommended);
    # "source" = prefer the source sheet's own coordinates where present.
    "coord_source": "db",

    # --- KPI units ----------------------------------------------
    # multiplier converting the UE-distance column into KILOMETRES
    # km -> 1.0 (current) | metres -> 0.001 | LTE TA steps (~78.07 m) -> 0.07807
    "ue_dist_to_km": 1.0,

    # --- flag thresholds ----------------------------------------
    "th_overshoot_idx": 1.20,
    "th_fair_share": 1.25,
    "th_isd_vs_median": 1.50,
    "th_att_share": 0.03,
    "min_att": 50,
}


# ============================================================
# COLUMN CANDIDATES
# Matched case- and punctuation-insensitively, then by substring,
# so "Avg UE distance (m)" or "AZ " still resolve. Add aliases freely.
# ============================================================

DB_COLS = {
    "cell":     ["cellName", "LNCEL name", "cell name"],
    "site":     ["Site name", "LNBTS name"],
    "sitecode": ["Site code", "sitecode"],
    "sec":      ["sec", "sec_id", "SEC"],
    "lat":      ["LAT", "latitude"],
    "lon":      ["LONG", "longitude"],
    "az":       ["AZ", "Azimuth"],
    "band":     ["band", "Band"],
    "earfcn":   ["earfcnDL", "earfcn"],
}

SRC_COLS = {
    "sitecode": ["Code", "Site code", "sitecode"],
    "sec":      ["SEC", "sec"],
    "lat":      ["latitude", "LAT"],
    "lon":      ["longitude", "LONG"],
    "az":       ["Azimuth", "AZ"],
    "band":     ["Band", "band"],
}

REL_COLS = {
    "src_site": ["Source LNBTS name"],
    "src_cell": ["Source LNCEL name"],
    "tgt_site": ["Target LNBTS name"],
    "tgt_cell": ["Target LNCEL name"],
    "att":      ["Number of Inter eNB Handover attempts per neighbor cell relationship",
                 "Inter eNB Handover attempts", "Handover attempts"],
}

# combined per-cell KPI export: UE distance, PRB, CQI, MCS, traffic, SE
# all identified by LNCEL/LNBTS name and read out of the same sheet.
PERCELL_COLS = {
    "cell":      ["LNCEL name", "cellName", "cell name"],
    "site":      ["LNBTS name", "Site name"],
    "dist":      ["Avg UE distance", "Average UE distance", "UE distance"],
    "pct":       ["% UEs with distance", "UEs with distance to base station"],
    "prb":       ["E-UTRAN Avg PRB usage per TTI DL", "Avg PRB usage per TTI DL",
                  "PRB usage per TTI DL"],
    "cqi":       ["Average CQI", "Avg CQI"],
    "mcs_pusch": ["Avg MCS PUSCH trans", "Avg MCS PUSCH"],
    "mcs_pdsch": ["Avg MCS PDSCH trans", "Avg MCS PDSCH"],
    "traffic":   ["PDCP SDU Volume", "PDCP SDU Volume, DL + UL", "PDCP Volume"],
    "se":        ["Spectral efficiency", "Spectral Efficiency", "Avg spectral efficiency"],
}

# how each per-cell KPI is aggregated when the export has multiple rows per
# LNCEL, and whether it is carried for the neighbour end too.
#   metric key -> (candidate key in PERCELL_COLS, "mean"|"sum", also_for_neighbour?)
PERCELL_KPIS = {
    "avg_ue_dist_km": ("dist",       "mean", True),
    "pct_ue_near":    ("pct",        "mean", True),
    "prb":            ("prb",        "mean", True),
    "cqi":            ("cqi",        "mean", True),
    "mcs_pusch":      ("mcs_pusch",  "mean", True),
    "mcs_pdsch":      ("mcs_pdsch",  "mean", True),
    "traffic":        ("traffic",    "sum",  True),
    "se":             ("se",         "mean", True),
}


# ============================================================
# IO helpers
# ============================================================

def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def pick(df, candidates, label="", required=True):
    """Resolve a column name from a candidate list."""
    lookup = {_norm(c): c for c in df.columns}

    for name in candidates:                        # exact, normalised
        if _norm(name) in lookup:
            return lookup[_norm(name)]

    for name in candidates:                        # substring, either direction
        k = _norm(name)
        if not k:
            continue
        for norm_col, original in lookup.items():
            if k in norm_col or norm_col in k:
                return original

    if required:
        raise KeyError(
            f"Could not find a column for '{label or candidates[0]}'.\n"
            f"  tried    : {candidates}\n"
            f"  available: {list(df.columns)}"
        )
    return None


def smart_read(path, sep=None, decimal=".", header_hint=None, label=""):
    """Read a CSV whose delimiter may be unknown and which may carry Nokia
    report preamble lines above the real header row."""
    if not path or not os.path.exists(path):
        print(f"  [skip] {label or path} not found.")
        return None

    with open(path, "r", encoding="latin1", errors="replace") as fh:
        head = [fh.readline() for _ in range(80)]

    skip = 0
    if header_hint:
        for i, line in enumerate(head):
            if line and any(h.lower() in line.lower() for h in header_hint):
                skip = i
                break

    line = head[skip] if skip < len(head) else ""
    if sep is None:
        counts = {c: line.count(c) for c in [";", "\t", ",", "|"]}
        sep = max(counts, key=counts.get)
        if counts[sep] == 0:
            sep = ","

    df = pd.read_csv(path, sep=sep, skiprows=skip, encoding="latin1",
                     engine="python", decimal=decimal, on_bad_lines="skip")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    print(f"  read {label or os.path.basename(path)}: {len(df)} rows, "
          f"sep={sep!r}, skiprows={skip}")
    return df


def to_num(series):
    """Numeric coercion surviving thousands separators, decimal commas,
    stray spaces and non-breaking spaces."""
    if series.dtype.kind in "if":
        return series
    s = (series.astype(str)
         .str.replace("\xa0", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.strip())
    cleaned = s.str.replace(",", ".", regex=False)
    both = s.str.contains(r"\.\d*,", regex=True, na=False)       # 1.234,56
    cleaned[both] = (s[both].str.replace(".", "", regex=False)
                            .str.replace(",", ".", regex=False))
    return pd.to_numeric(cleaned, errors="coerce")


def key(*parts):
    """Composite join key, whitespace- and case-insensitive."""
    out = None
    for p in parts:
        col = p.astype(str).str.strip().str.upper()
        out = col if out is None else out + "|" + col
    return out


# ============================================================
# Band handling
# ============================================================

BAND_ALIASES = {
    "L09": ["L09", "L9", "900", "B8", "LTE900"],
    "L18": ["L18", "1800", "B3", "LTE1800"],
    "L21": ["L21", "2100", "B1", "LTE2100"],
    "L08": ["L08", "800", "B20"],
    "TDD": ["TDD", "L23", "2300", "B40", "B41", "B38"],
}

EARFCN_RANGES = [
    (0, 599, "L21"), (1200, 1949, "L18"), (2750, 3449, "L26"),
    (3450, 3799, "L09"), (6150, 6449, "L08"),
    (37750, 38249, "TDD"), (38650, 39649, "TDD"), (39650, 41589, "TDD"),
]


def normalize_band(value):
    if pd.isna(value):
        return None
    v = _norm(value).upper()
    for label, aliases in BAND_ALIASES.items():
        if any(_norm(a).upper() == v for a in aliases):
            return label
    for label, aliases in BAND_ALIASES.items():
        if any(_norm(a).upper() in v for a in aliases):
            return label
    return str(value).strip().upper() or None


def band_from_earfcn(earfcn):
    if pd.isna(earfcn):
        return None
    try:
        e = int(float(earfcn))
    except (TypeError, ValueError):
        return None
    for lo, hi, label in EARFCN_RANGES:
        if lo <= e <= hi:
            return label
    return None


def resolve_band(df, band_col, earfcn_col=None):
    """EARFCN wins where available - it is unambiguous - with the band
    column as fallback so blank or oddly-spelled bands still resolve."""
    out = (df[band_col].map(normalize_band) if band_col
           else pd.Series(None, index=df.index, dtype=object))
    if earfcn_col and earfcn_col in df.columns:
        from_e = df[earfcn_col].map(band_from_earfcn)
        out = from_e.where(from_e.notna(), out)
    return out.fillna("UNK")


# ============================================================
# Geometry
# ============================================================

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(np.radians, [lat1, lon1, lat2, lon2])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2) ** 2
    return 2 * R * np.arcsin(np.sqrt(a))


def bearing(lat1, lon1, lat2, lon2):
    lat1, lat2 = map(np.radians, [lat1, lat2])
    dlon = np.radians(lon2 - lon1)
    x = np.sin(dlon) * np.cos(lat2)
    y = np.cos(lat1) * np.sin(lat2) - np.sin(lat1) * np.cos(lat2) * np.cos(dlon)
    return (np.degrees(np.arctan2(x, y)) + 360) % 360


def azimuth_diff(a, b):
    d = abs(a - b)
    return np.minimum(d, 360 - d)


# ============================================================
# Stage 0 - database and source resolution
# ============================================================

def load_sector_db(cfg):
    db = smart_read(cfg["sector_csv"], cfg["sector_sep"], cfg["decimal"],
                    header_hint=["cellName", "Site code"], label="sector DB")
    if db is None:
        raise FileNotFoundError(cfg["sector_csv"])

    c = {k: pick(db, v, label=f"DB.{k}", required=(k not in ("earfcn", "sec")))
         for k, v in DB_COLS.items()}

    out = pd.DataFrame({
        "Cell":     db[c["cell"]].astype(str).str.strip(),
        "Site":     db[c["site"]].astype(str).str.strip(),
        "SiteCode": db[c["sitecode"]].astype(str).str.strip(),
        "LAT":      to_num(db[c["lat"]]),
        "LON":      to_num(db[c["lon"]]),
        "AZ":       to_num(db[c["az"]]) % 360,
        "Band":     resolve_band(db, c["band"], c["earfcn"]),
    })
    out["Sec"] = db[c["sec"]].astype(str).str.strip().str.upper() if c["sec"] else ""

    bad = out["LAT"].isna() | out["LON"].isna() | out["AZ"].isna()
    if bad.any():
        print(f"  WARNING: dropping {bad.sum()} DB rows with missing LAT/LONG/AZ.")
        out = out[~bad]

    out["CellKey"] = key(out["Site"], out["Cell"])
    out["SectorKey"] = key(out["SiteCode"], out["Sec"], out["Band"])
    return out.reset_index(drop=True)


def resolve_source(cfg, db):
    """The source sheet is sector-level (Code / SEC / Band) with no cell name,
    so join it onto the DB to recover the LNCELs behind each sector."""
    src = smart_read(cfg["source_csv"], cfg["source_sep"], cfg["decimal"],
                     header_hint=["Code", "SEC"], label="source sheet")
    if src is None:
        raise FileNotFoundError(cfg["source_csv"])

    c = {k: pick(src, v, label=f"SRC.{k}", required=(k in ("sitecode", "sec")))
         for k, v in SRC_COLS.items()}

    s = pd.DataFrame({
        "SiteCode": src[c["sitecode"]].astype(str).str.strip(),
        "Sec":      src[c["sec"]].astype(str).str.strip().str.upper(),
    })
    s["Band"] = resolve_band(src, c["band"]) if c["band"] else "UNK"
    for name, col in (("LAT", c["lat"]), ("LON", c["lon"]), ("AZ", c["az"])):
        s[name] = to_num(src[col]) if col else np.nan
    s["SectorKey"] = key(s["SiteCode"], s["Sec"], s["Band"])

    merged = db.merge(s[["SectorKey"]].drop_duplicates(), on="SectorKey", how="inner")
    matched = set(merged["SectorKey"])

    leftover = s[~s["SectorKey"].isin(matched)]
    if len(leftover):                              # retry ignoring band
        loose_src = set(key(leftover["SiteCode"], leftover["Sec"]))
        db_loose = key(db["SiteCode"], db["Sec"])
        extra = db[db_loose.isin(loose_src) & ~db["SectorKey"].isin(matched)]
        if len(extra):
            print(f"  {leftover['SectorKey'].nunique()} source sectors did not match "
                  f"on band; {len(extra)} cells recovered by site+sector only.")
            merged = pd.concat([merged, extra], ignore_index=True)

    still = s[~s["SectorKey"].isin(set(merged["SectorKey"]))]
    if len(still):
        print(f"  WARNING: {still['SectorKey'].nunique()} source sectors found no cell "
              f"in the DB. Examples: {list(still['SectorKey'].unique()[:5])}")

    if cfg["coord_source"] == "source":
        coords = s.dropna(subset=["LAT", "LON"]).drop_duplicates("SectorKey")
        merged = merged.merge(
            coords[["SectorKey", "LAT", "LON", "AZ"]].rename(
                columns={"LAT": "sLAT", "LON": "sLON", "AZ": "sAZ"}),
            on="SectorKey", how="left")
        for a, b in (("LAT", "sLAT"), ("LON", "sLON"), ("AZ", "sAZ")):
            merged[a] = merged[b].where(merged[b].notna(), merged[a])
        merged = merged.drop(columns=["sLAT", "sLON", "sAZ"])

    merged = merged.drop_duplicates("CellKey").reset_index(drop=True)
    print(f"  resolved {len(merged)} source cells from "
          f"{s['SectorKey'].nunique()} source sectors.")
    return merged


# ============================================================
# Stage 1 - neighbour geometry
# ============================================================

def find_neighbors(src_cells, db, cfg):
    results = []

    for _, s in src_cells.iterrows():
        cand = db
        if cfg["same_band_only"]:
            cand = cand[cand["Band"] == s["Band"]]
        if cfg["exclude_cosite"]:
            cand = cand[cand["SiteCode"] != s["SiteCode"]]
        if cand.empty:
            continue

        cand = cand.copy()
        cand["ISD_km"] = haversine(cand["LAT"], cand["LON"], s["LAT"], s["LON"])
        cand = cand[cand["ISD_km"] > 0.001]
        if cfg["max_distance_km"]:
            cand = cand[cand["ISD_km"] <= cfg["max_distance_km"]]
        if cand.empty:
            continue

        b_nbr_src = bearing(cand["LAT"], cand["LON"], s["LAT"], s["LON"])
        b_src_nbr = bearing(s["LAT"], s["LON"], cand["LAT"], cand["LON"])
        cand["Nbr_AZ_Diff"] = azimuth_diff(cand["AZ"], b_nbr_src)
        cand["Src_AZ_Diff"] = azimuth_diff(s["AZ"], b_src_nbr)

        cand = cand[(cand["Src_AZ_Diff"] <= cfg["source_beamwidth"])
                    & (cand["Nbr_AZ_Diff"] <= cfg["nbr_beamwidth"])]
        if cand.empty:
            continue

        cand["Distance_Score"] = cand["ISD_km"] / cand["ISD_km"].max()
        cand["Alignment_Score"] = (cand["Src_AZ_Diff"] / cfg["source_beamwidth"]
                                   + cand["Nbr_AZ_Diff"] / cfg["nbr_beamwidth"]) / 2
        cand["Total_Score"] = (cfg["distance_weight"] * cand["Distance_Score"]
                               + cfg["azimuth_weight"] * cand["Alignment_Score"])

        cand = cand.sort_values("Total_Score").head(cfg["top_n"]).copy()
        cand = cand.rename(columns={
            "Cell": "Nbr_Cell", "Site": "Nbr_Site", "SiteCode": "Nbr_SiteCode",
            "Sec": "Nbr_Sec", "CellKey": "Nbr_Key",
        })

        cand["Src_Cell"] = s["Cell"]
        cand["Src_Site"] = s["Site"]
        cand["Src_SiteCode"] = s["SiteCode"]
        cand["Src_Sec"] = s["Sec"]
        cand["Src_Key"] = s["CellKey"]
        cand["Neighbor_Rank"] = range(1, len(cand) + 1)
        results.append(cand)

    return pd.concat(results, ignore_index=True) if results else pd.DataFrame()


# ============================================================
# Stage 2 - per-cell KPIs (UE distance, PRB, CQI, MCS, traffic, SE)
# All now read from a single combined export.
# ============================================================

def load_percell_kpi(path, sep, cfg, colmap, wanted, header_hint, label):
    """Load a per-cell KPI CSV and return one row per LNCEL with the wanted
    metrics.  `wanted` maps output_name -> (candidate_key_in_colmap, agg)."""
    kpi = smart_read(path, sep, cfg["decimal"], header_hint=header_hint, label=label)
    if kpi is None:
        return None

    c_cell = pick(kpi, colmap["cell"], label=f"{label}.cell")
    c_site = pick(kpi, colmap["site"], label=f"{label}.site", required=False)

    out = pd.DataFrame({"Cell": kpi[c_cell].astype(str).str.strip()})
    out["Site"] = kpi[c_site].astype(str).str.strip() if c_site else ""

    aggs = {}
    for name, (cand_key, agg) in wanted.items():
        col = pick(kpi, colmap[cand_key], label=f"{label}.{name}", required=False)
        if col is None:
            print(f"  NOTE: '{name}' column not found in {label} - skipped.")
            continue
        out[name] = to_num(kpi[col])
        aggs[name] = agg

    if not aggs:
        print(f"  WARNING: no usable KPI columns in {label}.")
        return None

    # collapse multi-row (per-day) exports to one row per cell
    grouped = out.groupby(["Site", "Cell"], as_index=False).agg(aggs)
    grouped["CellKey"] = key(grouped["Site"], grouped["Cell"])
    return grouped, list(aggs.keys())


def load_all_percell_kpi(cfg):
    """UE distance, PRB, CQI, MCS PUSCH/PDSCH, traffic and spectral efficiency
    now all live in one combined per-cell export - load and validate them
    together instead of as four separate files."""
    wanted = {name: (cand_key, agg) for name, (cand_key, agg, _) in PERCELL_KPIS.items()}
    result = load_percell_kpi(
        cfg.get("percell_kpi_csv"), cfg.get("percell_kpi_sep"), cfg, PERCELL_COLS,
        wanted=wanted,
        header_hint=["LNCEL name", "Avg UE distance", "Average CQI", "PRB usage"],
        label="per-cell KPI")
    if result is None:
        return None
    frame, metrics = result

    if "avg_ue_dist_km" in frame.columns:
        # apply the unit multiplier now - equivalent to applying it before
        # the mean aggregation above, since a constant factor commutes with mean
        frame["avg_ue_dist_km"] = frame["avg_ue_dist_km"] * cfg["ue_dist_to_km"]
        med = frame["avg_ue_dist_km"].median()
        print(f"  median avg UE distance = {med:.3f} km")
        if med > 20:
            print("  WARNING: implausibly large - the column is probably in metres. "
                  "Set ue_dist_to_km = 0.001.")
        elif med < 0.02:
            print("  WARNING: implausibly small - check ue_dist_to_km.")
    else:
        print("  NOTE: no UE-distance column found - overshoot_idx and "
              "fair_share_ratio cannot be computed.")

    return frame, metrics


def join_percell_kpi(rel, kpi, metrics, both_ends):
    """Attach per-cell metrics to the source end (always) and neighbour end
    (when listed in both_ends).  LNBTS+LNCEL first, unique-LNCEL-name fallback."""
    by_key = kpi.set_index("CellKey")[metrics]

    dup = kpi["Cell"].str.strip().str.upper().duplicated(keep=False)
    uniq = kpi[~dup].copy()
    uniq["k"] = uniq["Cell"].str.strip().str.upper()
    by_cell = uniq.set_index("k")[metrics]

    sides = [("src", "Src_Key", "Src_Cell")]
    sides += [("nbr", "Nbr_Key", "Nbr_Cell")]

    for side, keycol, namecol in sides:
        fk = rel[namecol].astype(str).str.strip().str.upper()
        for met in metrics:
            if side == "nbr" and met not in both_ends:
                continue
            joined = rel[keycol].map(by_key[met])
            joined = joined.where(joined.notna(), fk.map(by_cell[met]))
            rel[f"{met}_{side}"] = joined

    hit_col = f"{metrics[0]}_src"
    if hit_col in rel:
        print(f"  matched on {rel[hit_col].notna().mean():.0%} of relations "
              f"(source end, first metric).")
    return rel


# ============================================================
# Stage 3 - relation KPI (inter-eNB HO attempts - own independent export)
# ============================================================

def load_rel_kpi(cfg):
    rk = smart_read(cfg["rel_kpi_csv"], cfg["rel_sep"], cfg["decimal"],
                    header_hint=["Source LNCEL name"], label="neighbour HO KPI")
    if rk is None:
        return None

    c = {k: pick(rk, v, label=f"REL.{k}") for k, v in REL_COLS.items()}

    out = pd.DataFrame({
        "src_site": rk[c["src_site"]].astype(str).str.strip(),
        "src_cell": rk[c["src_cell"]].astype(str).str.strip(),
        "tgt_site": rk[c["tgt_site"]].astype(str).str.strip(),
        "tgt_cell": rk[c["tgt_cell"]].astype(str).str.strip(),
        "ho_att":   to_num(rk[c["att"]]),
    }).dropna(subset=["ho_att"])

    out["Src_Key"] = key(out["src_site"], out["src_cell"])
    out["Nbr_Key"] = key(out["tgt_site"], out["tgt_cell"])
    out["src_cell_u"] = out["src_cell"].str.upper()
    out["tgt_cell_u"] = out["tgt_cell"].str.upper()

    # a board export is usually multi-day: collapse to one row per relation
    return out.groupby(["Src_Key", "Nbr_Key", "src_cell_u", "tgt_cell_u"],
                       as_index=False)["ho_att"].sum()


def join_rel_kpi(rel, rk):
    """LNBTS+LNCEL on both ends first, then fill any gaps by cell-name pair.
    A missing match is usually a genuinely idle relation, not a naming problem,
    so the two are reported separately."""
    m = rel.merge(rk[["Src_Key", "Nbr_Key", "ho_att"]],
                  on=["Src_Key", "Nbr_Key"], how="left")
    strict = int(m["ho_att"].notna().sum())

    alt = rk.groupby(["src_cell_u", "tgt_cell_u"], as_index=False)["ho_att"].sum()
    alt = alt.rename(columns={"ho_att": "_alt"})
    m["_s"] = m["Src_Cell"].astype(str).str.strip().str.upper()
    m["_t"] = m["Nbr_Cell"].astype(str).str.strip().str.upper()
    m = m.merge(alt, left_on=["_s", "_t"], right_on=["src_cell_u", "tgt_cell_u"], how="left")
    m["ho_att"] = m["ho_att"].where(m["ho_att"].notna(), m["_alt"])
    m = m.drop(columns=["_s", "_t", "_alt", "src_cell_u", "tgt_cell_u"], errors="ignore")

    # was the shortfall caused by site naming, or by relations with no traffic?
    db_sites = set(rel["Src_Site"].astype(str).str.strip().str.upper())
    kpi_sites = set(rk["Src_Key"].str.split("|").str[0])
    overlap = len(db_sites & kpi_sites) / max(len(db_sites), 1)

    total = int(m["ho_att"].notna().sum())
    print(f"  {total} of {len(m)} relations matched an HO row "
          f"({strict} on LNBTS+LNCEL, {total - strict} on cell names).")
    if overlap < 0.5:
        print(f"  NOTE: only {overlap:.0%} of DB site names appear in the KPI export - "
              f"check 'Site name' vs 'LNBTS name' conventions.")
    unmatched = len(m) - total
    if unmatched:
        print(f"  {unmatched} relations have no HO row (idle relation, outside the "
              f"KPI scope, or missing from the neighbour list).")
    return m


# ============================================================
# Stage 4 - metrics
# ============================================================

def add_metrics(df, cfg):
    half = df["ISD_km"] / 2

    for col in ("avg_ue_dist_km_src", "avg_ue_dist_km_nbr", "ho_att"):
        if col not in df:
            df[col] = np.nan

    df["overshoot_idx_src"] = df["avg_ue_dist_km_src"] / half
    df["overshoot_idx_nbr"] = df["avg_ue_dist_km_nbr"] / half

    denom = (df["avg_ue_dist_km_src"] + df["avg_ue_dist_km_nbr"]).replace(0, np.nan)
    df["span_share_src"] = df["avg_ue_dist_km_src"] / denom      # 0.5 = fair
    df["dominance_point_km"] = df["ISD_km"] * df["span_share_src"]
    df["fair_share_ratio"] = 2 * df["span_share_src"]            # 1.0 = fair
    df["area_share_src"] = df["span_share_src"] ** 2 / (
        df["span_share_src"] ** 2 + (1 - df["span_share_src"]) ** 2)

    g = df.groupby("Src_Key")
    df["att_share"] = df["ho_att"] / g["ho_att"].transform("sum").replace(0, np.nan)

    def _wtd(x):
        tot = x["ho_att"].sum(skipna=True)
        return np.nan if not tot else (x["ho_att"] * x["ISD_km"]).sum(skipna=True) / tot

    df["att_wtd_nbr_dist_km"] = df["Src_Key"].map(g.apply(_wtd))
    df["isd_vs_median"] = df["ISD_km"] / g["ISD_km"].transform("median")

    rules = [
        df["overshoot_idx_src"] > cfg["th_overshoot_idx"],
        df["fair_share_ratio"] > cfg["th_fair_share"],
        df["isd_vs_median"] > cfg["th_isd_vs_median"],
        df["att_share"] > cfg["th_att_share"],
    ]
    df["rules_hit"] = sum(r.fillna(False).astype(int) for r in rules)

    enough = (df["ho_att"] >= cfg["min_att"]) | df["ho_att"].isna()
    df["severity"] = np.select(
        [~enough, df["rules_hit"] >= 3, df["rules_hit"] == 2, df["rules_hit"] == 1],
        ["LOW SAMPLE", "HIGH", "MEDIUM", "WATCH"], default="OK")
    df["overshoot_flag"] = df["severity"].isin(["HIGH", "MEDIUM"])
    return df


SEV_RANK = {"OK": 0, "LOW SAMPLE": 1, "WATCH": 2, "MEDIUM": 3, "HIGH": 4}
SEV_RANK_INV = {v: k for k, v in SEV_RANK.items()}


def cell_summary(df, percell_src_cols):
    agg = dict(
        n_relations=("Nbr_Cell", "count"),
        median_isd_km=("ISD_km", "median"),
        max_isd_km=("ISD_km", "max"),
        avg_ue_dist_km=("avg_ue_dist_km_src", "first"),
        max_overshoot_idx=("overshoot_idx_src", "max"),
        mean_overshoot_idx=("overshoot_idx_src", "mean"),
        max_fair_share_ratio=("fair_share_ratio", "max"),
        att_wtd_nbr_dist_km=("att_wtd_nbr_dist_km", "first"),
        total_att=("ho_att", lambda x: x.sum(min_count=1)),   # blank, not 0, if unmatched
        n_flagged=("overshoot_flag", "sum"),
    )
    # carry each source-side per-cell KPI (identical across a cell's rows -> first)
    for col in percell_src_cols:
        if col in df.columns:
            agg[col] = (col, "first")

    out = df.groupby(["Src_Key", "Src_Site", "Src_Cell", "Band"], as_index=False).agg(**agg)

    out["att_reach_ratio"] = out["att_wtd_nbr_dist_km"] / out["median_isd_km"]
    out["flagged_share"] = out["n_flagged"] / out["n_relations"]

    # worst-case verdict across the cell's relations, for one-click filtering
    worst = (df.assign(_r=df["severity"].map(SEV_RANK).fillna(0))
               .groupby("Src_Key")["_r"].max())
    out["cell_severity"] = out["Src_Key"].map(worst).map(SEV_RANK_INV)
    out["cell_overshoot_flag"] = out["cell_severity"].isin(["HIGH", "MEDIUM"])

    return out.sort_values(["cell_overshoot_flag", "n_flagged", "max_overshoot_idx"],
                           ascending=False)


# ============================================================
# Stage 5 - export
# ============================================================

def export(rel, cells, path):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    fills = {
        "HIGH":       PatternFill("solid", fgColor="FFC7CE"),
        "MEDIUM":     PatternFill("solid", fgColor="FFEB9C"),
        "WATCH":      PatternFill("solid", fgColor="FFF2CC"),
        "LOW SAMPLE": PatternFill("solid", fgColor="D9D9D9"),
    }

    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)

    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        rel.to_excel(xl, sheet_name="Relations", index=False)
        cells.to_excel(xl, sheet_name="Cell_Summary", index=False)

        for name, frame in (("Relations", rel), ("Cell_Summary", cells)):
            ws = xl.sheets[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for c in ws[1]:
                c.font = Font(bold=True)
            for i, col in enumerate(frame.columns, start=1):
                longest = frame[col].astype(str).str.len().max()
                longest = 12 if pd.isna(longest) else int(longest)
                ws.column_dimensions[get_column_letter(i)].width = max(12, min(30, longest + 2))

        # colour each sheet by its severity column
        for sheet, frame, sevcol in (("Relations", rel, "severity"),
                                     ("Cell_Summary", cells, "cell_severity")):
            if sevcol not in frame.columns:
                continue
            ws = xl.sheets[sheet]
            sev = list(frame.columns).index(sevcol) + 1
            for r in range(2, len(frame) + 2):
                fill = fills.get(ws.cell(row=r, column=sev).value)
                if fill:
                    for c in range(1, len(frame.columns) + 1):
                        ws.cell(row=r, column=c).fill = fill


# ============================================================
# Orchestration
# ============================================================

FRONT = [
    "Src_Site", "Src_Cell", "Src_Sec", "Nbr_Site", "Nbr_Cell", "Nbr_Sec",
    "Band", "Neighbor_Rank", "ISD_km", "Src_AZ_Diff", "Nbr_AZ_Diff",
    "avg_ue_dist_km_src", "avg_ue_dist_km_nbr",
    "overshoot_idx_src", "overshoot_idx_nbr",
    "dominance_point_km", "span_share_src", "fair_share_ratio", "area_share_src",
    "pct_ue_near_src", "ho_att", "att_share", "att_wtd_nbr_dist_km",
    "isd_vs_median",
    # per-cell load / quality / traffic KPIs (source then neighbour)
    "prb_src", "prb_nbr",
    "cqi_src", "cqi_nbr",
    "mcs_pusch_src", "mcs_pusch_nbr",
    "mcs_pdsch_src", "mcs_pdsch_nbr",
    "traffic_src", "traffic_nbr",
    "se_src", "se_nbr",
    "rules_hit", "severity", "overshoot_flag",
]

# order for the source-side per-cell KPIs on the Cell_Summary sheet
PERCELL_SRC_ORDER = ["prb_src", "cqi_src", "mcs_pusch_src", "mcs_pdsch_src", "traffic_src", "se_src"]


def main(cfg=CONFIG):
    print("Stage 0: database and source resolution ...")
    db = load_sector_db(cfg)
    print(f"  bands present: {db['Band'].value_counts().to_dict()}")
    src_cells = resolve_source(cfg, db)
    if src_cells.empty:
        print("No source cells resolved - check that Code/SEC match "
              "'Site code'/'sec' in the DB.")
        return None

    print("Stage 1: geometry ...")
    rel = find_neighbors(src_cells, db, cfg)
    if rel.empty:
        print("No neighbours found. Loosen beamwidths or max_distance_km.")
        return None
    print(f"  {len(rel)} relations across {rel['Src_Key'].nunique()} source cells.")

    print("Stage 2: per-cell KPI (UE distance / PRB / CQI / MCS / traffic / SE) ...")
    percell = load_all_percell_kpi(cfg)
    if percell is not None:
        kpi_frame, metrics = percell
        both_ends = {name for name, (_, _, both) in PERCELL_KPIS.items() if both}
        rel = join_percell_kpi(rel, kpi_frame, metrics, both_ends)

    print("Stage 3: neighbour HO KPI ...")
    rk = load_rel_kpi(cfg)
    if rk is not None:
        rel = join_rel_kpi(rel, rk)

    print("Stage 4: metrics ...")
    rel = add_metrics(rel, cfg)
    cells = cell_summary(rel, PERCELL_SRC_ORDER)

    order = [c for c in FRONT if c in rel.columns]
    rel = rel[order + [c for c in rel.columns if c not in order]]
    for frame in (rel, cells):
        num = frame.select_dtypes(include=[np.number]).columns
        frame[num] = frame[num].round(4)

    print("Stage 5: export ...")
    export(rel, cells, cfg["out_xlsx"])

    print("\nSeverity breakdown:")
    print(rel["severity"].value_counts().to_string())
    print(f"\nDone -> {cfg['out_xlsx']}")
    return rel, cells


if __name__ == "__main__":
    main()
