"""
Deep validator for a Purview audit export against the M365 dashboard's needs.

Streams the CSV (handles multi-GB files), tallies per-Operation and per-RecordType
counts, samples AuditData JSON for the nested fields each dashboard page consumes,
and reports gaps that would cause visuals/DAX/M to fail or render empty.
"""

from __future__ import annotations
import csv, json, sys, io
from collections import Counter, defaultdict
from pathlib import Path

csv.field_size_limit(min(sys.maxsize, 2**31 - 1))

CSV_PATH = Path(sys.argv[1] if len(sys.argv) > 1
                else r"C:\Users\luzlorenz\Downloads\79a33305-4ce1-477f-afd8-4ca83afc0991.csv")

REQUIRED_COLS = ["RecordId", "CreationDate", "RecordType", "Operation", "UserId", "AuditData"]

CANONICAL_OPS = {
    "MailItemsAccessed", "MailboxLogin", "Send",
    "FileAccessed", "FileViewed", "FilePreviewed", "FileModified", "FileDownloaded", "FileUploaded",
    "MessageSent", "MessageRead", "MessagesListed", "ChatRetrieved", "ChatCreated",
    "MeetingParticipantJoined", "MeetingStarted", "MeetingEnded",
    "MeetingParticipantDetail", "MeetingDetail",
    "TeamsSessionStarted",
    "CopilotInteraction", "ConnectedAIAppInteraction",
}

# Required AuditData JSON keys per Operation family (sampled from a few rows each).
# Empty set means "no nested-field requirement beyond Operation + UserId + CreationTime".
NEEDLES = {
    "CopilotInteraction": {"CopilotEventData", "AppHost"},
    "ConnectedAIAppInteraction": {"AppHost"},  # agent fields live in AppHost / CopilotEventData
    "MeetingDetail": {"MeetingDetail"},
    "MeetingParticipantDetail": {"Attendees"},
    "MailItemsAccessed": {"OperationProperties"},
    "FileAccessed": {"SiteUrl"},
}

SAMPLE_PER_OP = 5  # parse AuditData for first N rows of each op

def main():
    if not CSV_PATH.exists():
        print(f"FILE NOT FOUND: {CSV_PATH}")
        return 2

    size_gb = CSV_PATH.stat().st_size / (1024 ** 3)
    print(f"File           : {CSV_PATH}")
    print(f"Size           : {size_gb:.2f} GB")

    op_counts = Counter()
    rec_counts = Counter()
    workload_counts = Counter()
    users_per_op: dict[str, set] = defaultdict(set)
    samples: dict[str, list[dict]] = defaultdict(list)
    needle_hits: dict[str, Counter] = defaultdict(Counter)
    min_dt = max_dt = None
    bad_rows = 0
    total = 0
    header = None

    # Open with utf-8-sig to strip BOM if present.
    with CSV_PATH.open("r", encoding="utf-8-sig", newline="", errors="replace") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        missing_cols = [c for c in REQUIRED_COLS if c not in header]
        idx = {c: header.index(c) for c in header}

        for row in reader:
            total += 1
            try:
                op = row[idx["Operation"]]
                rt = row[idx["RecordType"]]
                uid = row[idx["UserId"]]
                cd = row[idx["CreationDate"]].strip()
                ad = row[idx["AuditData"]] if "AuditData" in idx else ""
            except (IndexError, KeyError):
                bad_rows += 1
                continue

            op_counts[op] += 1
            rec_counts[rt] += 1
            if uid:
                users_per_op[op].add(uid)
            if cd:
                if min_dt is None or cd < min_dt:
                    min_dt = cd
                if max_dt is None or cd > max_dt:
                    max_dt = cd

            if op in NEEDLES and len(samples[op]) < SAMPLE_PER_OP:
                try:
                    j = json.loads(ad)
                    samples[op].append(j)
                    wl = j.get("Workload")
                    if wl:
                        workload_counts[wl] += 1
                    for needle in NEEDLES[op]:
                        if needle in j:
                            needle_hits[op][needle] += 1
                except (json.JSONDecodeError, ValueError):
                    pass
            elif op in ("CopilotInteraction", "ConnectedAIAppInteraction") and total % 5000 == 0:
                # extra workload sampling for AI ops
                try:
                    wl = json.loads(ad).get("Workload")
                    if wl:
                        workload_counts[wl] += 1
                except Exception:
                    pass

            if total % 500_000 == 0:
                print(f"  ... scanned {total:>10,} rows", file=sys.stderr)

    print("\n" + "=" * 72)
    print("HEADER / SCHEMA")
    print("=" * 72)
    print(f"Columns        : {header}")
    if missing_cols:
        print(f"MISSING COLS   : {missing_cols}  <-- BLOCKER")
    else:
        print("Columns OK     : all 6 required columns present")
    print(f"Total rows     : {total:,}")
    print(f"Malformed rows : {bad_rows}")
    print(f"Date range     : {min_dt}  ->  {max_dt}")

    print("\n" + "=" * 72)
    print("OPERATION COVERAGE vs CANONICAL 22-OP LIST")
    print("=" * 72)
    present = set(op_counts) & CANONICAL_OPS
    missing = CANONICAL_OPS - set(op_counts)
    extra = set(op_counts) - CANONICAL_OPS

    print(f"\nPresent ({len(present)}/{len(CANONICAL_OPS)}):")
    for op in sorted(CANONICAL_OPS):
        if op in op_counts:
            n = op_counts[op]
            u = len(users_per_op[op])
            flag = "" if n > 0 else "  <-- ZERO ROWS"
            print(f"  [OK]   {op:<32} rows={n:>10,}  users={u:>6,}{flag}")
        else:
            print(f"  [MISS] {op:<32}                     <-- MISSING")

    if missing:
        print(f"\n*** {len(missing)} REQUIRED OPS MISSING: {sorted(missing)}")
    else:
        print("\nAll 22 required operations are present in the export.")

    if extra:
        print(f"\nExtra ops in export (harmless, not used by dashboard): {len(extra)}")
        top_extra = sorted(((op_counts[o], o) for o in extra), reverse=True)[:15]
        for n, o in top_extra:
            print(f"  - {o:<40} rows={n:,}")

    print("\n" + "=" * 72)
    print("RECORDTYPE DISTRIBUTION (top 15)")
    print("=" * 72)
    for rt, n in rec_counts.most_common(15):
        print(f"  RecordType={rt:<4}  rows={n:>10,}")

    print("\n" + "=" * 72)
    print("AUDITDATA NESTED-FIELD CHECK (sampled)")
    print("=" * 72)
    for op, needles in NEEDLES.items():
        if op not in op_counts:
            print(f"  {op}: op absent, skipped")
            continue
        sampled = len(samples[op])
        print(f"\n  {op}  (sampled {sampled} rows)")
        for n in sorted(needles):
            hit = needle_hits[op].get(n, 0)
            ok = "OK " if hit > 0 else "!! "
            print(f"    {ok}{n:<30} hits={hit}/{sampled}")

    # Agent-telemetry deep check for AI ops
    print("\n" + "=" * 72)
    print("AGENT TELEMETRY DEEP CHECK (Copilot License Optimizer)")
    print("=" * 72)
    for op in ("CopilotInteraction", "ConnectedAIAppInteraction"):
        if op not in op_counts:
            print(f"  {op}: not present, CLO page WILL render empty for this slice")
            continue
        print(f"\n  {op}: {op_counts[op]:,} rows, {len(users_per_op[op]):,} users")
        # walk samples for agent fields
        agent_keys_seen = Counter()
        for j in samples[op]:
            ah = j.get("AppHost", "")
            ced = j.get("CopilotEventData", {}) or {}
            if isinstance(ced, dict):
                for k in ced.keys():
                    agent_keys_seen[k] += 1
            if ah:
                agent_keys_seen[f"AppHost={ah}"] += 1
        if agent_keys_seen:
            print(f"    AppHost / CopilotEventData keys observed in samples:")
            for k, n in agent_keys_seen.most_common(15):
                print(f"      - {k}  ({n})")
        else:
            print("    !! No AppHost / CopilotEventData seen in sampled rows")

    print("\n" + "=" * 72)
    print("WORKLOAD DISTRIBUTION (from sampled AuditData)")
    print("=" * 72)
    for wl, n in workload_counts.most_common(15):
        print(f"  {wl:<30} samples={n}")

    print("\n" + "=" * 72)
    print("VERDICT")
    print("=" * 72)
    blockers = []
    warnings = []
    if missing_cols:
        blockers.append(f"Missing required columns: {missing_cols}")
    if missing:
        blockers.append(f"Missing required Operations: {sorted(missing)}")
    zero_rows = [op for op in present if op_counts[op] == 0]
    if zero_rows:
        warnings.append(f"Ops present in header but zero rows: {zero_rows}")
    # CLO sanity
    if op_counts.get("ConnectedAIAppInteraction", 0) == 0:
        warnings.append("ConnectedAIAppInteraction has 0 rows -> CLO 'Connected AI apps' page will be empty.")
    if op_counts.get("CopilotInteraction", 0) == 0:
        warnings.append("CopilotInteraction has 0 rows -> Copilot adoption page will be empty.")
    # Meeting completeness
    meeting_ops = ("MeetingDetail", "MeetingParticipantDetail", "MeetingStarted", "MeetingEnded")
    if not any(op_counts.get(o, 0) for o in meeting_ops):
        warnings.append("No meeting-family ops present -> Teams Meetings page will be empty.")

    if blockers:
        print("BLOCKERS:")
        for b in blockers:
            print(f"  - {b}")
    if warnings:
        print("WARNINGS:")
        for w in warnings:
            print(f"  - {w}")
    if not blockers and not warnings:
        print("GREEN. File has every Operation, RecordType, and nested AuditData field "
              "the dashboard reads. Safe to feed into build_rollup_with_agents.py.")
    elif not blockers:
        print("YELLOW. No blockers, but check warnings above before running the pipeline.")
    else:
        print("RED. Fix blockers before running the pipeline.")

    return 0

if __name__ == "__main__":
    sys.exit(main())
