# Derivative Cover Check

_Owner: Hilton Netta | Team: Investment Operations / Compliance | Frequency: Daily (business days)_

> This documents an **existing, already-built** pipeline (`derv_checker.py` and its stages), not a tool to be built from scratch. The Process section below was reconstructed from reading the code; sections marked **[CONFIRM]** are things only Hilton knows and need to be filled in or corrected.

---

## Context

Every business day, each fund's derivative positions need enough cash cover behind them to stay within the derivative cover limit set out in Chapter V of Board Notice 90. This check pulls each fund's holdings and derivative exposures from Eagle, calculates cash cover for every fund, and produces two distinct outputs for two distinct audiences:

- **`Derv {date}.xlsx`** — used only by **Compliance** (Hilton), to (i) get an overall view of derivative cover across all funds, and (ii) generate each fund's individual derivative cover check sheet, e.g. `PIMBAL Derv Check {date}.xlsx`.
- **`Free Cover.xlsm`** — emailed daily to the Prescient Investment Management investment teams (`pimequity@prescient.co.za`, `cashteam@prescient.co.za`, `fixedinterest@prescient.co.za`, `multiasset@prescient.co.za`) with subject line `Free Derivative Cover - {day date month year}` (e.g. "Free Derivative Cover - Thursday 13 August 2026"). This tells the investment teams how much free cash each of their funds has before breaching the Board Notice 90 Chapter V cover limit.

**What happens if this doesn't run, or runs with wrong numbers:** the investment teams, not being informed (or being informed with inaccurate data), could trade a fund into breach of the Board Notice 90 Chapter V limit — i.e. aggregate effective exposure from listed and OTC derivatives exceeding the fund's assets in liquid form ("cash").

**Deadline:** there's no house-set deadline, but the file is built on fund holdings and values as at the prior business day's close, so it loses its value to the investment teams if it hasn't been emailed to them by roughly noon.

---

## Outputs

| File | Location | What it is |
|---|---|---|
| `{fund} Derv Calc {ddMMMyyyy}.xlsx` | `...\Compliance\Derivative Cover\` | One workbook per fund: holdings + derivative deltas + that fund's cash cover calc |
| `Derv {ddMMMyyyy}.xlsx` | `...\Compliance\Derivative Cover\` | Combined summary across all funds, sorted by "Cash Cover for UT" ascending (worst cover first), grouped by UT type (UT / ≠UT / UCITS / SAA / TAA / ICAV) |
| `{yyyymmdd}_derv_calc.xlsx`, `{yyyymmdd}_holdings.xlsx` | `...\Compliance\Derivative Cover\` (holdings one under `\Holdings\`) | Combined holdings/derivatives tables built by `derv_checker_table.py`, separate from the per-fund files above |
| `Free Cover.xlsm` (updated, not created) | `...\Compliance\Daily\` | The master compliance file. Also used as the pipeline's own "already ran today" check (see step 3 below) |

**Frequency:** Daily, business days only (`prior_working_day()` skips weekends/SA public holidays).

**What the 10% (`dervthreshold`) actually flags:** it's not a hard breach line — the Board Notice 90 Chapter V breach itself is at 0% cover (effective exposure exceeding liquid assets). `dervthreshold` marks funds that are "at risk": within 10% of their cash cover being exceeded by derivative effective exposure. `Free Cover.xlsm`, emailed to the investment teams, highlights a summary of these at-risk funds specifically, and includes a hyperlink to the full free cover summary for all PIM-managed funds at `P:\Investment Operations\GRC\Compliance\Derivative Cover\Derv {date}.xlsx` for anyone who wants the complete picture, not just the at-risk subset.

---

## Inputs

### Input 1: Control settings — `py_reports.xlsm`, sheet `arc`

- **Source:** `...\Compliance\Daily\py_reports.xlsm`, read by `parn_de()` in `utilities.py`
- **Column A:** the fund code list to run (one per row, from row 2 down)
- **Column E, fixed rows** (independent of how many funds are listed):
  - Row 2: report date override (if blank, defaults to the prior working day)
  - Row 4: `summ_yn` — whether a full summary sheet run is required
  - Row 5: `dervthreshold` — the cover threshold %, e.g. currently `0.1` → 10%
  - Row 6: `batches` — number of batches used when downloading from Eagle for a large fund list

### Input 2: Eagle "Portfolio Analytics Report - New" (holdings) — `PARN`

- **Source system:** Eagle portal (`eagleportal.prescient.co.za`), pulled via Selenium (`osprey()` in `utilities.py`, report type `"parn"`)
- **How it's fetched:** logs into Eagle, opens the Portfolio Analytics report, sets the date and fund list, exports and downloads as CSV
- **Format:** CSV, saved to `Downloads\` as `PARN ({n funds}) {ddMMMyyyy}.csv`
- **Key columns:** `Entity Name`, `Entity ID`, `PrimaryAssetID`, `Maturity Date`, `Current Exposure`, `Sum of Market Value Income`, `% of Total Market Value`

### Input 3: Eagle "Derivative Exposure" report — `DERV`

- **Source system:** same as above, report type `"derv"`
- **Format:** CSV, saved as `DERV ({n funds}) {ddMMMyyyy}.csv`
- **Key columns:** `Entity Name`, `Primary Asset ID`, `Nominal Holding`, `Delta`, `Market Value`, `Effective Exposure`

### Input 4: Fund / UT lookup

- **Source:** `pthSttlmnt` = `P:\Investment Operations\GRC\Compliance\Daily\fund_codes.xlsx`, sheet `Funds` (fund code, fund name, UT status, investment team; the code here reads columns A:B, D:E off this sheet)
- This workbook also has a `Sttlmnt` tab (each fund's settlement bank) and a `fund` tab (fund long names, manager names, applicable regulations, etc.) — used extensively across the other `.py` fund reporting scripts, not just this one

### Input 5: Per-fund calc template

- **Source:** `derv_tmpl` (a template workbook `derv_checker_compiling.py` copies and fills in per fund)

---

## Process

1. **`derv_checker.py` → `derv_check()`** is the entry point.
2. **Get today's settings:** `parn_de()` reads the `arc` sheet (Input 1) → fund list, report date, `summ_yn`, `dervthreshold`, `batches`.
3. **Skip-if-already-done check:** reads `Free Cover.xlsm`'s `Summary` sheet, cell C1, for a date. If it already matches today's report date, the whole pipeline is skipped (today's run is assumed already complete) and the script just logs that fact.
4. Otherwise, runs these stages in order, each a separate script via `subprocess`:
   1. **`derv_checker_downloading.py`** — downloads Inputs 2 and 3 (PARN, DERV) from Eagle for every fund in the list.
   2. **`derv_checker_table.py`** — loads the two CSVs, cleans/casts numeric and date columns, recalculates `% of Total Market Value` and `Current Exposure %` per fund, relabels "UNKNOWN" TRS swaps, merges holdings with derivative deltas (on `Entity Name` + `PrimaryAssetID`), and writes the combined `{yyyymmdd}_derv_calc.xlsx` / `{yyyymmdd}_holdings.xlsx` files. Also appends bank overdraft balances where available.
   3. **`derv_checker_compiling.py`** — for each fund not already built today (checked against a local temp cache folder), builds that fund's individual `{fund} Derv Calc {ddMMMyyyy}.xlsx` from the template, saves it locally, then copies all of today's local files to the network `Derivative Cover` folder and clears the local cache.
   4. **`derv_checker_summarising.py`** — reads every fund's per-fund workbook `Summary` sheet and collates them into the single `Derv {ddMMMyyyy}.xlsx`, sorted by cash cover (worst first).
   5. **`derv_checker_freecover.py`** — writes today's results into `Free Cover.xlsm` (the file step 3 checks on the next run).
   6. **`derv_checker_cact.py`** — if both the PARN and DERV downloads exist, imports cash activity/flows for the funds (runs regardless of the skip-check in step 3).

**How it's actually run:** the intention and the ideal is to run the entire pipeline through once, end-to-end, via `derv_check()`. In practice there are occasions (not ideal) where manual intervention is needed to re-run one stage script by hand.

**Stage naming:** each stage's variable name in `constants.py` follows a `dc_` + two-letter pattern:
- `dc_do` — Derivative Cover DOwnload (`derv_checker_downloading.py`)
- `dc_tb` — Derivative Cover TaBle (`derv_checker_table.py`)
- `dc_co` — Derivative Cover COmpiling (`derv_checker_compiling.py`)
- `dc_su` — Derivative Cover SUmmarising (`derv_checker_summarising.py`)
- `dc_fr` — Derivative Cover FRee cover (`derv_checker_freecover.py`)

**Work in progress:** Hilton is currently working to collapse the pipeline from `dc_do → dc_tb → dc_co → dc_su → dc_fr` down to `dc_do → dc_tb → dc_fc` (`dc_fc` = Derivative Cover Free Cover, `derv_checker_freec.py` — currently present in `derv_checker.py` but commented out as "NEW freecover sheet"). Not finished yet, so the live pipeline for now is still the five-stage one above.

---

## Verification

Hilton is fairly comfortable with the accuracy of the outputs, given the pipeline logic itself. Anomalies almost always trace back to the Eagle source data from `dc_do`, not the calculations — most commonly, call/put option deltas and derivative effective exposure values missing from the `Delta` and `Effective Exposure` columns of the `DERV ({n funds}) {date}.csv` file.

**This can't be fixed by Compliance.** When it happens, Compliance raises the missing values with the Instruments team (currently headed by Sean Scorer) at Prescient Fund Services ("PFS") — PIM's fund accounting service provider and the custodian of Eagle — so they can insert the missing values directly in Eagle.

**Where this is caught in code:** `derv_checker_table.py` builds a `funds_missing_exposure` dataframe — derivatives with a non-zero `Nominal Holding`, not a SYTH, but a missing or zero `Effective Exposure` — and writes it to the `deltas_missing` sheet of `{yyyymmdd}_derv_calc.xlsx`, then prints the affected funds and tickers to the console (around lines 983–1018).

---

## Notes from recent debugging (13–14 Aug 2026)

- `parn_de()` in `utilities.py` had a row-alignment bug that dropped the settings rows (report date / `summ_yn` / `dervthreshold` / `batches`) whenever the fund list was short — fixed.
- `derv_checker_table.py` and `derv_checker_compiling.py` both compare `Maturity Date` (datetime64) against `rptDate + 397 days`; since `rptDate` is now a plain `date`, that comparison needed wrapping in `pd.Timestamp(...)` — fixed in both files.
- A stale, truncated per-fund file left in the local temp cache from an earlier interrupted run caused a `BadZipFile` error in `derv_checker_summarising.py` — resolved by re-running `derv_checker_compiling.py`, which regenerates any fund not already validly present.
