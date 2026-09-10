
# Identify (and optionally delete) old-format "Derv Calc" files, e.g.
# "ECICBALC Derv Calc 31Aug2026.xlsx", in the Derivative Cover folder

import os
import re
from constants import pthEXPORTS

pattern = re.compile(r"^[A-Za-z0-9_]{4,10} Derv Calc \d{2}[A-Za-z]{3}\d{4}\.xlsx$")

matches = sorted(
    f
    for f in os.listdir(pthEXPORTS)
    if os.path.isfile(os.path.join(pthEXPORTS, f)) and pattern.match(f)
)

print(
    f"{len(matches)} matching file{'' if len(matches) == 1 else 's'} "
    f"in {pthEXPORTS}:\n"
)
for f in matches:
    print(f"  {f}")

if not matches:
    raise SystemExit

confirm = (
    input(f"\nDelete all {len(matches)} file(s) listed above? [y/N] ")
    .strip()
    .lower()
)
if confirm == "y":
    for f in matches:
        os.remove(os.path.join(pthEXPORTS, f))
        print(f"  deleted {f}")
else:
    print("No files deleted.")
