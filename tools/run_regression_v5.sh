#!/usr/bin/env bash
# v.5 regression runner: Slicer/PivotTable shards from _regshards/*.txt manifests.
# Guardrails (memory rules 12/14):
#   - filters built from the authoritative --list-tests manifest (exact FQN contains-match)
#   - a shard without a "已通过!/失败!" summary line counts as FAILED (zero-match / crash trap)
#   - executed total must equal the shard's `.count` (contains-semantics case count)
#   - transient COM (0x800706BA / RPC) failures get exactly one retry after cleanup + rest
#   - every attempt keeps its own log (`_rr_v5_<shard>_a<N>.log`) so failure evidence survives
#
# Usage:
#   bash tools/run_regression_v5.sh                 # all shards
#   bash tools/run_regression_v5.sh D_olap          # one shard
set -u
# Derive the workspace root from this script's own location; override with EXCEL_MCP_ROOT when the
# tree is checked out elsewhere. Everything below uses paths relative to that root.
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="${EXCEL_MCP_ROOT:-$(cd "$HERE/.." && pwd)}"
cd "$ROOT"

REPORT=_regression_v5_report.txt
: > "$REPORT"
SHARDS=("$@")
if [ ${#SHARDS[@]} -eq 0 ]; then
    SHARDS=(A_lifecycle B_fields C_layout_grouping D_olap E_slicers_charts)
fi

parse_summary() {  # -> "failed passed skipped total"
    sed -E 's/.*失败: *([0-9]+)，通过: *([0-9]+)，已跳过: *([0-9]+)，总计: *([0-9]+).*/\1 \2 \3 \4/'
}

for s in "${SHARDS[@]}"; do
    mf="_regshards/$s.txt"
    cnt="_regshards/$s.count"
    if [ -f "$cnt" ]; then
        expected=$(tr -d ' \r\n' < "$cnt")
    else
        expected=$(sort -u "$mf" | grep -c .)
    fi
    filter=$(sort -u "$mf" | awk '{printf "%sFullyQualifiedName~%s", sep, $0; sep="|"}')
    verdict="UNKNOWN"
    for attempt in 1 2; do
        log="_rr_v5_${s}_a${attempt}.log"
        echo "=== $s attempt$attempt expected=$expected start=$(date +%H:%M:%S)" | tee -a "$REPORT"
        bash build.sh test --no-build -c Release \
            tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj \
            --filter "$filter" > "$log" 2>&1
        rc=$?
        line=$(grep -m1 -E "已通过!|失败!" "$log")
        if [ -z "$line" ]; then
            echo "  NO-SUMMARY rc=$rc log=$log (crash or zero-match)" | tee -a "$REPORT"
            verdict="NO-SUMMARY"
            if [ "$attempt" -eq 1 ]; then
                MSYS_NO_PATHCONV=1 taskkill /F /IM EXCEL.EXE >/dev/null 2>&1
                sleep 60
                continue
            fi
            break
        fi
        read -r f p sk tot <<< "$(echo "$line" | parse_summary)"
        echo "  rc=$rc failed=$f passed=$p skipped=$sk total=$tot end=$(date +%H:%M:%S) log=$log" | tee -a "$REPORT"
        if [ "$tot" -ne "$expected" ]; then
            echo "  MISMATCH: total($tot) != expected($expected)" | tee -a "$REPORT"
        fi
        if [ "$f" -gt 0 ] && grep -qE "0x800706BA|RPC 服务器不可用" "$log"; then
            echo "  transient COM failure detected (see $log)" | tee -a "$REPORT"
            verdict="TRANSIENT"
            if [ "$attempt" -eq 1 ]; then
                MSYS_NO_PATHCONV=1 taskkill /F /IM EXCEL.EXE >/dev/null 2>&1
                sleep 60
                continue
            fi
        elif [ "$f" -gt 0 ]; then
            verdict="FAILED($f)"
        else
            verdict="PASS"
        fi
        break
    done
    echo "=== $s VERDICT: $verdict" | tee -a "$REPORT"
    sleep 45
done

echo "ALL SHARDS DONE $(date +%H:%M:%S)" | tee -a "$REPORT"
