"""Generate regression shard manifests for the v.5 Slicer/PivotTable run.

Source of truth: _tlist_v5.log (`dotnet test --list-tests`, built 2026-09-15 16:22).
Each shard file: one test-case line per entry (from the authoritative list).
Filters use `FullyQualifiedName~<ClassFqn>.<MethodPrefix>` (contains match,
no parentheses - parameterized test args would break VSTest filter syntax).
"""

import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LIST_LOG = ROOT / "_tlist_v5.log"
OUT_DIR = ROOT / "_regshards"
OUT_DIR.mkdir(exist_ok=True)

# (shard, method-prefix regex) - first match wins
RULES = [
    ("A_lifecycle", r"^(Create(FromRange|FromTable|FromDataModel|ScenarioSummary)|Delete_|List_|GetData_|GetInfo_|DataPreparation|DrillThrough|CacheOptions|Refresh_|PivotTableDispatch|RemoveField_|SetFieldName_|CreateCalculatedField|DeleteCalculatedField|ListCalculatedFields|SetFieldFormat|SetFieldFunction)"),
    ("B_fields", r"^(Add(Column|Filter|Row|Value)Field_|SortField_|SetField(Name|Function)_)"),
    ("C_layout_grouping", r"^(SetLayout_|SetSubtotals_|SetGrandTotals_|Group(By|Items)|SetTableSlicerSelection|CreateTableSlicer|DeleteTableSlicer|ListTableSlicers)"),
    ("D_olap", r"(Olap|DataModel|CalculatedMember|CubeFields|Measures?|OlapPivot)",),
    ("E_slicers_charts", r".*"),  # catch-all: Slicer*, SlicerLink_*, CreateSlicer_*, CreateFromPivotTable_*, rest
]

tests = []          # (args-stripped fqn, raw line) - in-scope cases, one entry per list line
universe = []       # every list line (args-stripped): what a filter can actually match
for line in LIST_LOG.read_text(encoding="utf-8", errors="replace").splitlines():
    t = line.strip()
    if not re.match(r"^    \S", line):
        continue
    fqn = t.split("(")[0].strip()  # strip parameterized args (VSTest filter cannot take parens)
    universe.append(fqn)
    if not ("Slicer" in t or "PivotTable" in t):
        continue  # out of scope for the manifest, but still reachable by a filter
    tests.append((fqn, t))

shards = defaultdict(list)
raw_counts = defaultdict(int)
for fqn, raw in tests:
    method = fqn.rsplit(".", 1)[-1]
    for name, pattern in RULES:
        if re.match(pattern, method) or re.search(pattern, method):
            shards[name].append(fqn)
            raw_counts[name] += 1
            break

# Expected counts follow the filter's CONTAINS semantics: a list line executes if its
# args-stripped FQN contains some manifest entry. This is how the parameterized pair
# (`...(reportType: Summary)` / `...(reportType: PivotTable)`) yields 2 cases from 1 entry.
all_stripped = universe
total = 0
for name in ["A_lifecycle", "B_fields", "C_layout_grouping", "D_olap", "E_slicers_charts"]:
    items = shards.get(name, [])
    entries = sorted(set(items))
    expected = sum(1 for fqn in all_stripped if any(e in fqn for e in entries))
    total += expected
    p = OUT_DIR / f"{name}.txt"
    p.write_text("\n".join(items) + "\n", encoding="utf-8")
    (OUT_DIR / f"{name}.count").write_text(str(expected) + "\n", encoding="utf-8")
    print(f"{name:<20} manifest_entries={len(entries):3d} expected_cases={expected:3d}  -> {p.name}")
print("TOTAL expected cases", total, "of", len(tests), "list lines")
