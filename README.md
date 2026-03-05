# IRP Cleanup

Processes Integrated Resource Plan (IRP) data from Minnesota's four largest utilities into tidy CSV files for nameplate capacity and energy mix projections.

**Note**: IRP spreadsheets are obtained from Steve Rakow at the Minnesota Department of Commerce. Analysts will need to reach out to Steve directly to refresh
the data. Utilities submit IRP data at non-standard intervals, so I recommend that analysts should reach out to Steve roughly every 6 months or so. The spreadsheet
does get changed frequently, so analysts should expect to have to update this cleaning script with every new IRP sheet. 

## Source Data

The script reads from a pre-formatted Excel workbook (`raw_data/IRP_Cleanup_3-2-2026 modified.xlsx`) containing one tab per utility — Xcel, GRE, MP, and OTP — each with a Nameplate (MW) table and an Energy Mix (GWh) table. A Summary tab aggregates across utilities. The workbook is derived from IRP filings and standardized so that all four utility tabs share the same row structure.

## Setup

You'll need to create two directories in the project root before running:

```
raw_data/
processed_data/
```

Both are in `.gitignore` and won't be tracked. Place the source Excel workbook in `raw_data/`.

### Dependencies

- `tidyverse`
- `tidyxl`
- `janitor`
- `here`

## What the Script Does

**Nameplate capacity** (`capacity_cleaner`): Extracts projected nameplate capacity in MW from rows 1–17 of each utility tab. Filters out non-generation categories (contracts, demand response, energy efficiency, etc.), consolidates the three gas/oil sub-types into a single "Natural Gas" category, and renames fuel types to plain labels (e.g., "Coal:Conventional" → "Coal"). Adds end-of-horizon labels for charting.

**Energy mix** (`mix_cleaner`): Extracts projected generation in GWh from rows 20–36 of each utility tab. Applies the same filtering, consolidation, and renaming logic as the capacity cleaner.

Both functions are applied across all four utilities and combined with `bind_rows`.

## Outputs

- `processed_data/irp_nameplate_projections.csv` — long-format nameplate capacity by fuel type, year, and utility
- `processed_data/irp_energy_mix_projections.csv` — long-format generation by fuel type, year, and utility

## Notes

- If the source workbook structure changes (different row ranges, new tabs), update the row filter ranges in both `capacity_cleaner` and `mix_cleaner`. Comments in the code flag these lines.
- The `write_csv` paths currently point to a shared network drive (`I:/...`). Update these to use `here("processed_data/...")` if running outside that environment.