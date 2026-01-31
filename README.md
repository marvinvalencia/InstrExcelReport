# INSTR Logger → Excel Report Converter

A lightweight Windows application that converts **INSTR Data Logger CSV exports** into a formatted Excel report with automatic calculations, summaries, and charts.

**This tool is for Tito Ian.**

---

## What this tool does

- Reads INSTR logger CSV exports (UTF-16, tab- or comma-delimited)
- Parses metadata, channel definitions, and logged readings
- Down-samples data to full elapsed minutes
- Calculates temperature rise (ΔT) from ambient
- Generates a ready-to-use Excel report with:
  - Raw data
  - Automatic summaries
  - Charts
  - Configurable thermocouple groupings
  - Observations sheet for test notes

---

## Output

The generated Excel report includes:

- **Summary of Results** with charts
- **Raw Data** with absolute values and ΔT calculations
- **Config** sheet for thermocouple grouping adjustments
- **Observations** sheet for free-form notes

---

## Usage

1. Launch the application
2. Select the INSTR CSV file
3. Review or adjust thermocouple grouping options
4. Click **Generate Report**
5. The Excel report is saved in the same folder as the input CSV

---

## Notes

- Both tab- and comma-delimited INSTR exports are supported
- Thermocouple groupings can be adjusted after generation via the **Config** sheet
- Charts and formulas update automatically when config values change
