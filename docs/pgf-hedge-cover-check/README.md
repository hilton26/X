# PGF Hedge Cover Check

_Owner: Hilton Netta | Team: Investment Operations / Compliance | Frequency: Daily (business days), reported by T+1_

> This documents an **existing, already-built** pipeline (`pgf_checker.py` → `pgf_downloading.py` → `pgf_compiling.py`), not a tool to be built from scratch. The Process section below was reconstructed from reading the code; sections marked **[CONFIRM]** are things only Hilton knows and need to be filled in or corrected. There's also a **Possible issues found while documenting** section at the bottom — two things that look like bugs, flagged for you to confirm before I touch any code.

---

## Context

PGF (Prescient Global Funds plc) is a UCITS umbrella domiciled in Ireland, some of whose share classes are currency-hedged. This check pulls each hedged share class's unit trust price/NAV and holdings, and produces a daily "Share Class Hedges" workbook showing the extent to which each PGF UCITS hedge class sits within the currency hedging limits prescribed by **Regulation 26 of the Central Bank UCITS Regulations** (the Central Bank of Ireland, PGF's regulator).

**The 95%–105% thresholds are regulatory, not informational.** A hedge class outside that band must be adhered to or remedied immediately on identification — this isn't a "nice to know," it's a live regulatory limit.

**Who relies on it:** emailed daily to the investment teams:
- **To:** `MultiAsset@prescient.co.za`, `PIMEquity@prescient.co.za`, `fixedinterest@prescient.co.za`, `pimcashandincome@prescient.co.za`
- **Cc:** `pfsimancorisk@prescient.ie`, `Nasreen.Hisham@prescient.co.za`, `nazley.herandien@prescient.co.za`, `richard.flint@prescient.co.za`, `riskanalytics@prescient.co.za`
- **Subject line:** `PGF Share Class Hedging - {date} - {"Not all" or "All"} hedges are within the regulatory limits` (i.e. the subject itself states the breach/no-breach outcome)

**What happens if this is skipped:** the investment teams won't know the extent of their UCITS class currency hedging, and risk trading blind — potentially exacerbating an existing breach. Beyond that, it risks the wrath and ire of the depositary (Northern Trust plc) and the regulator (the Central Bank of Ireland), with penalties charged against the scheme manager, Prescient Global Funds plc.

**Deadline:** performed and reported — breach or no breach — every business day by `T+1`.

---

## Outputs

| File | Location | What it is |
|---|---|---|
| `{yyyymmdd} PGF Share Class Hedges.xlsx` | `...\Compliance\PGF UCITS Share Class Hedges\` | That day's workbook: `Portfolio Valuation` (holdings), `Class NAVs` (unit trust prices), `Summary` (with the prior day's values pasted in for comparison) |
| `PGF Share Class Hedges.xlsm` | same folder | The template file itself, re-saved after being populated (acts as the latest snapshot / working copy) |
| Daily email (manual, not yet scripted) | sent to the investment teams — see Context | Subject states the breach/no-breach outcome directly; body is a copy-paste of cells `A1:J16` of the `Summary` sheet (via a VBA script), not an attachment |

---

## Inputs

### Input 1: Control settings — `py_reports.xlsm`, sheet `arc`, columns G:I

- **Source:** same control file as the derivative cover check (`pthPy`), but a different section of the same `arc` sheet
- **Column G** (`pgf: UT prices`): PGF share class codes to pull unit trust prices/NAVs for, e.g. `PGPCEMD`, `PGPCGED`, `PGPGBFD`, `PGPGIFE`, `PGPGIFI`, `PGPRFE`, `PGPRFG`
- **Column H** (`pgf: PAR-N`): PGF share class codes to pull holdings for, e.g. `PGPRUSD`, `PGPRZAR`, `PGPRUS0%`, `PGPRGBP`, `PGIPZAR`, `PGIP_I`, `PGPCZAR`, `PGBEUR`, `PGBGBP`, `PGBZAR`, `PGEMZAR`
- **Column I** (`pgf_checker`), row 1 (first data row): report date override — if blank, defaults to the prior working day. **See "Possible issues" below — the code currently reads the wrong row for this.**

### Input 2: Eagle "Unit Trust Prices" report — `UTPS`

- **Source system:** Eagle, via `osprey()`, report type `"utps"`
- **Format:** CSV, `UTPS PGF_UT_Prices({n classes}) {ddMMMyyyy}.csv` in `Downloads\`
- **Key columns:** `NAV Price`, `Clean Price`, `Income Price`, `Class Size`, `Units in Class`, `Effective Date`

### Input 3: Eagle "Portfolio Analytics Report - New" (holdings) — `PARN`

- **Source system:** Eagle, via `osprey()`, report type `"parn"`
- **Format:** CSV, `PARN PGF_Holdings({n classes}) {ddMMMyyyy}.csv` in `Downloads\`
- **Key columns:** `Original Nominal`, `Sum of Market Value Income`, `% of Total Market Value`, `Current Exposure`, `Current Exposure %`, `i Position Effective Date`

### Input 4: Prior day's workbook

- **Source:** the most recent existing `{yyyymmdd} PGF Share Class Hedges.xlsx` already in `...\Compliance\PGF UCITS Share Class Hedges\` (found by taking the max of the 8-digit date prefixes present) — its `Summary` sheet rows 1–11, column I, are pasted into today's `Summary` sheet column J for day-over-day comparison

### Input 5: Template workbook

- **Source:** `pth_hdg_tmpl` = `...\Compliance\PGF UCITS Share Class Hedges\PGF Share Class Hedges.xlsm`, opened via `xlwings`, sheets `Portfolio Valuation`, `Class NAVs`, `Summary`

---

## Process

1. **`pgf_checker.py` → `pgf_check()`** is the entry point.
2. Reads Input 1 (`arc` sheet, columns G:I) for the report date, and the two share-class lists.
3. **Skip-if-already-done check:** if today's `{yyyymmdd} PGF Share Class Hedges.xlsx` already exists in the hedges folder, the whole pipeline is skipped and the script just logs that fact.
4. Otherwise, runs two stages via `subprocess`:
   1. **`pgf_downloading.py`** — downloads Inputs 2 and 3 (UTPS, PARN) from Eagle for the two share-class lists, skipping a download if its file already exists locally. **See "Possible issues" below.**
   2. **`pgf_compiling.py`** — runs only if both download files exist. Loads the two CSVs, casts numeric/date columns, pulls in the prior day's `Summary` values (Input 4), pastes holdings/NAVs/prior-day-comparison into the template (Input 5) via `xlwings`, and saves it both as today's dated file and back over the template `.xlsm` itself. Finally moves the two downloaded CSVs into `pthTest` (an archive/scratch folder) so they don't linger in `Downloads\`.
5. **Manual step (not yet scripted in this repo):** Compliance reviews the compiled workbook (`PGF Share Class Hedges.xlsm`, and the current/prior dated copies alongside it, both in `...\Compliance\PGF UCITS Share Class Hedges\`) against the 95%–105% Regulation 26 thresholds, then sends the daily email to the investment teams. A VBA script copy-pastes cells `A1:J16` of the `Summary` sheet directly into the email body (no attachment), and the subject line states whether all hedges are within the regulatory limits.

**[CONFIRM]**
- Do you ever run `pgf_downloading.py` / `pgf_compiling.py` individually, or always the full `pgf_check()`?

---

## Verification

**Sanity check:** confirm that every active hedge class has data obtained for the prior day's close from Eagle. Short of that, PFS must be contacted to confirm whether PGF UCITS fund accounting was completed and received for the day — a missing class most often means the fund accounting cycle at PFS hasn't finished yet, not a data-pull problem on this end.

**[CONFIRM]** Anything else worth noting here — e.g. a typical/expected number of active hedge classes, what a "normal" day-over-day comparison (Summary column I vs J) looks like vs. something clearly wrong, or anything that's gone wrong before and what the actual cause was.

---

## Issues found while documenting (both fixed)

1. **Report date override was being read from the wrong row.** `pgf_checker.py`, `pgf_downloading.py`, and `pgf_compiling.py` were reading the date override as `df.iloc[1, 2]` (second data row of column I), but the live `arc` sheet has the override value in the **first** data row. Fixed by Hilton — all three now read `df.iloc[0, 2]`, confirmed against the live sheet to correctly resolve the override date instead of silently falling back to `prior_working_day(datetime.today())`.
2. **The "already downloaded" skip-checks in `pgf_downloading.py` were swapped.** Before downloading UT prices (`osprey("utps", ...)`), it was checking whether the *holdings* file (`parN_nm`) already existed, and vice versa for the holdings download. Fixed: each block now checks the existence of the file it's actually about to fetch (`utP_nm` before the UTPS download, `parN_nm` before the PARN download).
