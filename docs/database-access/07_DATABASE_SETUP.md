# 07 — Connecting to PIM Databases
_The optional step. Only do this if your work needs live data from a PIM database._

> **Claude, read this whole file before doing anything.** Your job is to get this
> person connected to PIM databases safely, the read-only way. Work in this order,
> one step at a time, and confirm each step worked before moving on:
> 1. Check Python 3.12 and `pip` are available (Step 2).
> 2. Install `ppym` using the private index command in Step 3 exactly as written.
> 3. Append the "Database access rules" block from Step 4 to the `CLAUDE.md` in
>    their workspace. Do not remove anything already in their `CLAUDE.md` — add to it.
> 4. Run the small test in Step 5 and show them the result.
> If anything errors, fix it before continuing. Never use any credential other than
> `read.only` / `read.only` unless the person explicitly tells you to for a specific task.

---

## Do you actually need this?

Most first automations work on files you already have: an Excel export, a CSV, a
folder of PDFs. If that is your task, skip this file. You do not need a database
connection to automate a report you build from a download.

Do this step when your task needs to **pull data straight from a PIM system** rather
than from a file you saved. For example: fund returns, prices, holdings, ASISA
categories, benchmark series. If you find yourself logging into Prime and exporting
the same query every day, this is for you.

---

## The one rule that keeps everyone safe

**Always connect with the read-only user.**

```
username: read.only
password: read.only
host:     pim-cpt-mysql-prod.prescient.local
port:     3306  (main databases)
          3307  (EAV / time-series: prime_eav, prime_profile, pim_ts, pim_static, benchmarks)
```

This account can only read. It cannot insert, update, delete, drop, or change
anything. Our databases are live production systems with no separate test copy, so
the same database you query also runs client reporting and portfolio management.
Read-only means Claude cannot break anything, no matter what it tries. This is the
standard, and it was signed off by IT Security (Calvyn) for exactly this purpose.

You need to be on the Prescient network or VPN for the connection to reach the host.

---

## Step 1 — Python on PATH

You need Python 3.12. IT installs it for you, but check one thing: during install,
the box **"Add Python to PATH"** must be ticked, or the terminal cannot find it.

Confirm it works. Open a terminal in VS Code (`Ctrl+` `) and run:

```bash
python --version
```

You want `Python 3.12.x`. If you see "not recognized", Python is installed but not on
PATH. Tell Claude what you see and it will help, or message IT to re-run the installer
with the PATH box ticked.

---

## Step 2 — Install ppym (the PIM Python Module)

`ppym` is the shared library that knows how to talk to our databases. It handles the
connections, the entity lookups, and the right way to fetch a time series, so you do
not have to. It lives on a private Prescient index, not the public internet, so the
install command points pip at that index:

```bash
pip install --extra-index-url https://docs-ppym.prescient.co.za/downloads/ ppym --force-reinstall
```

That command is the one thing Claude cannot guess on its own, which is why it is
written here. The `--extra-index-url` tells pip where to find ppym; `--force-reinstall`
makes sure you get the current build.

Confirm it installed:

```bash
python -c "import ppym; print('ppym ready')"
```

- **Docs:** https://docs-ppym.prescient.co.za/ppym.html
- **Source and README:** https://gitlab-pim.prescient.co.za/dsqa/pim-python-module

> If you are one of the few people who develop ppym itself (you have the repo cloned
> and edit it directly), do **not** pip install over your working copy. That note is
> for the maintainers; everyone else should install with the command above.

Three packages are required for the connection tests in Step 4. Install them now:

```bash
pip install sqlalchemy pymysql bizdays
```

`sqlalchemy` and `pymysql` power the direct connection test. `bizdays` is a ppym
internal dependency that is not bundled with ppym itself — you will hit a
`ModuleNotFoundError` in Step 4 without it.

Any other package you need (pandas, openpyxl, and so on) Claude can find and install
with plain `pip install` in your session. ppym is the only one with a special source.

---

## Step 3 — Teach your Claude how to connect

So that Claude connects the safe way in every future session, add the rules to the
`CLAUDE.md` in your workspace (the file Claude reads automatically at the start of
each conversation). Ask Claude:

> "Add these database access rules to my CLAUDE.md. Don't remove anything already there, just add this section."

Then paste:

```markdown
## Database access rules

- ALWAYS connect with `username='read.only'`, `password='read.only'`.
- Host: `pim-cpt-mysql-prod.prescient.local`
- Port 3306 for main databases; port 3307 for EAV / time-series (prime_eav, prime_profile, pim_ts, pim_static, benchmarks).
- NEVER use a credential with write access unless I explicitly tell you to for one specific, approved task.
- Prefer `ppym` for connections and queries. Connection patterns and worked examples:
  ppym source — https://gitlab-pim.prescient.co.za/dsqa/pim-python-module (ppym/data/db.py)
  ppym docs   — https://docs-ppym.prescient.co.za/ppym.html
- Add a `LIMIT` to exploratory queries (start with `LIMIT 100`).
- Avoid `SELECT *` on wide tables; name the columns you need.
- Don't run queries in a tight loop.
- I am the human in the loop: show me the SQL before you run it.
```

Once that is in your `CLAUDE.md`, you never have to re-explain any of it.

> **Where these rules really live.** The block above is the simplified on-ramp. The
> maintained, version-controlled source is the `pim-claude` repo (see `08`), kept
> current by Tim (DSQA) and the PRA team and approved by IT Security. Once you have
> GitLab access, you can have Claude pull the latest version straight from the repo
> instead of this copy. This kit mirrors that source; it is not a second authority.

---

## Step 4 — Connect and test

Two ways to connect. Use the ppym way for real work; the direct way is a quick check
that everything is wired up.

**The ppym way (recommended):**

```python
import os
from ppym.data.db import create_engine_multi, get_series

# read-only credentials
db = create_engine_multi(["prime_eav"], "read.only", "read.only")

# a small fetch (an example daily price series item)
df = get_series(
    list_entities=["NDDUWI Index"],
    list_items=["RT116"],
    list_source_codes=["BB"],
    db_engine=db["prime_eav"],
)
print(df.tail())

[db[x].dispose() for x in db]   # always close engines when done
```

**The direct way (a quick connection test, no ppym needed):**

```python
# needs: pip install sqlalchemy pymysql bizdays
from sqlalchemy import create_engine, text

engine = create_engine(
    "mysql+pymysql://read.only:read.only@pim-cpt-mysql-prod.prescient.local:3307/prime_profile"
)
with engine.connect() as conn:
    rows = conn.execute(text("SELECT COUNT(*) FROM tmp_funds")).scalar()
    print("rows in tmp_funds:", rows)
```

If either prints a result, you are connected.

---

## What is in the databases

A short map so you know what to ask for. The full catalogue, with table-by-table
detail, lives in the PRA knowledge base (`Knowledge/PIM-Data/DATABASE_GUIDE.md`) —
ask the PRA team or point your Claude at it.

| Database | Port | What it holds |
|---|---|---|
| `prime_eav` | 3307 | The central time-series store: returns, prices, ratios, flows, fund size. Access via ppym (`get_series`). |
| `prime_profile` | 3307 | SA and global fund register, ASISA categories, fund metrics. `tmp_funds` is the live universe. |
| `prime_eagle` | 3306 | The internal Prescient fund system mirror: the single source of truth for our own fund returns, instruments, holdings. |
| `prime_credit` | 3306 | Credit and bond static data, plus the canonical instrument classification. |
| `prime_ma` | 3306 | Multi-asset portfolio metadata. |
| `benchmarks`, `pim_ts`, `pim_static` | 3307 | Benchmark series and static reference data. |

The full list of databases on each port is in the access-rules block above and in the
PRA database guide.

---

## POPIA and data safety

Same rule as the rest of this kit. Claude sends your prompts and any data you show it
to servers outside South Africa. So:

- **Safe to query and discuss:** structure, column names, public market data, ASISA
  categories, published fund information, aggregate numbers.
- **Not safe to send to Claude:** client names, account numbers, individual positions,
  unpublished holdings, personal information.

The pattern that keeps you safe: let Claude write the query and the tool, then run it
yourself so the sensitive rows stay on your machine. When in doubt, ask the PRA team.

---

## When it does not work

- **"can't connect" / timeout** → You are probably off the network. Connect to the
  Prescient VPN and try again.
- **SSL or certificate error on install** → That is the Netskope certificate. The fix
  and a copy of the certificate are in `01_AI_SETUP` and the "Netskope Combined
  Certificate" folder in this kit. Restarting VS Code after setting it usually clears it.
- **"Access denied for user 'read.only'"** → Double-check the spelling: it really is
  `read.only` for both the username and the password.
- **ppym won't import after install** → Confirm Python 3.12 with `python --version`,
  then re-run the install command in Step 2.
- **Still stuck** → Screenshot it and message the PRA team, or bring it to First Friday.

---

## Getting help

| Channel | When |
|---|---|
| **First Friday 14:00-15:00, Zambezi** | PRA team and AI Champions, hands-on |
| **Teams message** | Ruan, Keasha, or Sophy. Send a screenshot of the error and which step you are on. |
| **The PIM AI Champions Teams channel** | Shared questions, shared answers |

---

_Built for the PIM AI Empowerment programme by Ruan Yacumakis and the PRA team.
The read-only access standard is maintained with Tim (DSQA) and approved by IT Security._
