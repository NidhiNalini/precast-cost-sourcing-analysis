# Changelog

All notable changes to this project are recorded here.
Format based on [Keep a Changelog](https://keepachangelog.com/); versions follow [Semantic Versioning](https://semver.org/).

## [1.1.0] - 2026-10-24
### Changed
- Lookups now use `XLOOKUP` with its *if not found* argument instead of `IFERROR(INDEX/MATCH)`
- Taller header rows so long column titles are fully visible
- Wider columns on the Sourcing and Summary sheets; wrapped headers on Summary
- Gridlines hidden on report sheets for a cleaner layout

### Added
- `--compat` option in `build_workbook.py` to build an `INDEX/MATCH` version for older Excel
- Screenshots of the Cost_Rollup, Sourcing and Data_Quality sheets in the README

### Removed
- Editor history folder that was committed by mistake (now in `.gitignore`)

## [1.0.0] - 2026-10-21
### Added
- Synthetic data generator (`scripts/generate_data.py`) with deliberate data quality issues
- Workbook builder (`scripts/build_workbook.py`)
- `Precast_Cost_Analysis.xlsx` with cost roll-up, sourcing, summary, data quality and open questions sheets
- Power Query (M) code for components, BOM and sales
- Exercises guide and README
