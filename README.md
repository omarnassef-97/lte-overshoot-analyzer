# lte-overshoot-analyzer

Geometric neighbor-relation and KPI-driven overshoot detection for LTE networks.

## Problem

In a dense LTE network, cells can end up "overshooting" — radiating well beyond their intended coverage footprint and creating unwanted dominance over neighboring cells. Overshooting cells are a frequent root cause of poor handover performance, interference, and degraded KPIs, but spotting them by hand across thousands of sectors is slow and inconsistent.

This tool automates that detection by combining the geometry of each cell's neighbor relations with the KPI data that shows whether a given relation is actually causing harm.

## Approach

- **Geometry first.** For every source cell, the script resolves its candidate neighbors from a sector database, then scores each relation using haversine distance and azimuth/beamwidth alignment to flag pairs that are geometrically implausible for healthy coexistence (too far apart relative to inter-site distance, poorly aligned antennas, cross-band mismatches, etc.).
- **KPIs confirm impact.** Per-cell KPI exports and neighbor handover-analysis exports are joined onto the geometry, so a flagged relation is only reported as an overshoot candidate when there's a real traffic/attempt share behind it, not just a geometric oddity.
- **Resilient to messy exports.** Column names in real-world network-management exports vary release to release and vendor to vendor. The script resolves each expected field from a list of known aliases per column instead of assuming a fixed header, and falls back gracefully when a column is missing.
- **Configurable thresholds.** Overshoot index, fair-share ratio, inter-site-distance ratio, and minimum attempt counts are all adjustable constants, so the same script can be retuned for a different network or vendor without touching the logic.
- **Ranked output.** Results are exported to Excel, ranked so the top-N most impactful overshooting relations per cell are easy to action first.

## Tech stack

Python · pandas · NumPy · openpyxl

## How to run

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Drop your exports into `data/`:
   - `source.csv` — source cell list
   - `sector_db.csv` — full sector/cell database (coordinates, azimuth, beamwidth, band)
   - `percell_kpi.csv` — per-cell KPI export (e.g. average MCS/PUSCH throughput)
   - `neighbor_ho.csv` — neighbor handover-analysis export
3. Adjust the `CONFIG` dict at the top of `main.py` if your file names, delimiters, or thresholds differ.
4. Run:
   ```
   python main.py
   ```
5. The ranked overshoot report is written to `output/overshoot_analysis.xlsx`.

## Project history

This is the fourth iteration of the script. Earlier versions grew four separate, near-duplicate code paths for loading and joining per-cell KPI data as new KPI sources were added one at a time. This version consolidates all of that into a single configurable loading/joining path, making it straightforward to add another KPI source without copy-pasting a new block.

## Note on data

No real network data is included. The `data/` and `output/` folders are gitignored — the script expects you to supply your own exports in the formats described above.
