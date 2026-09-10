# Compliance Breach Review

_Owner: Hilton Netta | Backup: Nazley Herandien (Compliance Manager) | Team: Investment Operations / Compliance | Frequency: Daily (business days), two runs_

> **Unlike the derivative cover check, this routine is currently mostly manual.** `portal.py` and `breaches.ipynb` are early, unfinished attempts at automating part of it — not a working pipeline yet. This README documents the current manual process as the source of truth, and the automation attempts separately, so it's clear what's real today vs. in progress.

---

## Context

The Compliance system runs an investment-limit breach check for all PIM funds, based on each fund's holdings and valuations as at the prior business day's close (`T`), and reports the result the following day (`T+1`).

**Who relies on it:** the daily breach distribution list — **not identical** to the derivative cover check's list:

- **To:** `MultiAsset@prescient.co.za`, `PIMEquity@prescient.co.za`, `fixedinterest@prescient.co.za`, `pimcashandincome@prescient.co.za`
- **Cc:** `pfsimancorisk@prescient.ie`, `riskanalytics@prescient.co.za`, `Nasreen.Hisham@prescient.co.za`, `nazley.herandien@prescient.co.za`, `richard.flint@prescient.co.za`

**What happens if this is skipped or late:**
1. The investment teams can only act to clear a breach after `T+1` days — later than intended.
2. Clients would only learn of their fund's breach after `T+1` days, which breaches Prescient's covenant with clients that the investment-limit breach system detects and reports breaches within `T+1` days.

---

## Process (current — manual)

There are **two Compliance runs per day**, both against fund holdings/valuations as at the prior business day's close:

1. **~11:11am run** — an investment-limit review across all funds, including fund groups whose fund accounting/pricing hasn't closed yet for that day's cycle:
   - **SA-domiciled funds** — already priced/closed by this time.
   - **Ireland-domiciled UCITS funds under the PGF scheme** — not yet priced; their NAVs are struck around 1pm SA time. Funds: QIFFGIF, PCGEARF, PGPGIF_C, PGPRF_C, PGPGBF_C, PGPCGE_C, PGPGARF, PGPCEM, PGCEF, PGCBF.

   Because the PGF UCITS prices aren't struck yet at 11:11am, any AM breaches attributable to the PGF funds themselves — or to SA-domiciled funds holding PGF UCITS funds — can safely be disregarded at this stage.

2. **~2:11pm run** — once PGF UCITS prices have been struck, breaches are reviewed in full and reported to the investment teams.

**Review steps, each run:**
1. An automated email arrives from `pim-no-reply@prescient.co.za`, subject `PIM Compliance Local Run - {date}`, once the run completes.
2. This prompts the Compliance officer (Hilton, or the backup when out of office) to open the Portal: `https://prime.prescient.co.za/portfolio-management/compliance/breach-report-new`
3. Each reported breach is considered for validity by the Compliance officer.
4. Each valid breach is reported by email to the team managing the affected fund — the recipient is looked up via `pthSttlmnt` (`fund_codes.xlsx`), sheet `Funds`, using the `Fund Code` and `PIM Investment Team Email` columns.

---

## Sanity check

- Expect roughly **20–40 breaches** at the 11:11am run.
- Of those, the ones traceable to the not-yet-priced PGF UCITS funds (or SA funds holding them) are the ones safe to set aside pending the 2:11pm run — they aren't real breaches yet, just an artefact of the PGF funds not having struck a price.

- A check on how frequently the same breach has occurred in the same fund over the recent past (three months) — a recurring breach points to something structural in that fund rather than a one-off.

---

## Automation attempts (in progress, not finished)

- **`portal.py`** — Selenium script; an early attempt at automating this. Its stated objective is to extract the breach record from `https://prime.prescient.co.za/portfolio-management/compliance/breach-report-new` and produce a summary of open breaches split into:
  - **Current** — occurred in the past working day
  - **Open** — occurred more than one working day ago, still unresolved

  What it does today: logs into the Portal, and for breach types `New` / `Open` / `Unclassified` (excluding `Resolved`), pulls each as CSV and writes it as a sheet into one dated workbook, `...\Compliance\Breaches\Portal\{timestamp} PortalBreachReg.xlsx`. It does **not** yet produce the Current/Open split described above — that's the gap.

- **`breaches.ipynb`** — another early attempt at the same objective. Reads a separate `breaches_register_cumulative.xlsx` and builds a fund × rule breach-count heatmap. How data is meant to flow from `PortalBreachReg.xlsx` into that cumulative register is **still unresolved** — noted by Hilton as something to work out, possibly with help, later.

---

## Related, but out of scope for this README

- **`r28_ib.py`** compiles the **monthly** Pension Funds Act Regulation 28 / Medical Schemes Act Regulation 30 instrument classification report (Schedule IB), as part of a separate pipeline: look-through download → merge → classify (`r_classifier`) → Reg 28 Schedule IB (`r28_ib.py`) and Reg 28 Table 2 Infrastructure report (`r28_t2.py`). This is **fully separate** from the daily breach review documented here.
- Several similarly-named files exist in `src/` for that pipeline (`lt_dl.py`, `lt_dl_and_merge.py`, `lt_dl_merge.py`, `lt_merge.py`, `lt_1_by_1.py`, `r28_t2.py`, `r28_tbl2.py`, `r28_Tbl2_bulk.py`, etc.) — worth its own README later to pin down which are current vs. draft/superseded.
