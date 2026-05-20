"""
build_rollup_combined.py
────────────────────────────────────────────────────────────────────────────
Combine N Purview audit CSV exports into ONE rollup CSV that the existing
PBIT consumes unchanged.

Why N files: the validated pull strategy is 4 separate Purview exports
(Teams, Outlook, Files, Copilot) — each scoped tightly so none hit the
1M-row cap. This script streams all of them through one aggregation
pass so the PBIT still sees a single rollup file with the same schema.

Output schema (identical to build_rollup_with_agents.py — required by the
M365Usage.tmdl column-fingerprint validator):

   UserId, CreationDate, Operation, Workload, SourceFileExtension,
   AppHost, EventCount, CreationTime, MaxCreationTime,
   AgentId, AgentName, ContextType, IsAgentInteraction

Operation handling:
 - KEEP_OPERATIONS reflects the 14 Operation literals actually referenced
   by the DAX model (verified by audit_dax_operations.py).
 - OP_RENAME normalises legacy / mislabelled names emitted by older
   Purview exports OR pasted by users into the wrong filter:
       FileViewed                  -> FileAccessed
       MeetingParticipantJoined    -> MeetingParticipantDetail
       ConnectedAIAppInteraction   -> AIAppInteraction
 - CANONICAL_OPS supplies case-insensitive normalisation.

Usage:
   python build_rollup_combined.py <output.csv> <input1.csv> [input2.csv ...]

Example (the validated 4-pull set):
   python build_rollup_combined.py rollup.csv ^
       teams.csv outlook.csv files.csv copilot.csv

Future: PAX ingestion. A PAX export is expected to land in the same
output schema directly (no AuditData JSON parsing required). When that
arrives, point the PBIT at either this rollup OR the PAX file — the
column fingerprint is the same so the model loads either source.
"""
import csv
import json
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

csv.field_size_limit(10 * 1024 * 1024)

# ──────────────────────────────────────────────────────────────────────
# Operation name handling
# ──────────────────────────────────────────────────────────────────────

# The 14 Operations actually referenced by the DAX model.
KEEP_OPERATIONS = {
    "CopilotInteraction", "AIAppInteraction",
    "FileAccessed", "FileModified", "FileDownloaded", "FileUploaded",
    "MailItemsAccessed", "Send", "MailboxLogin",
    "MessageSent", "MessageRead", "ChatCreated",
    "TeamsSessionStarted", "MeetingParticipantDetail",
}

# Legacy / mislabelled names that customers may have in historical
# exports. Rewritten to the canonical name BEFORE the KEEP filter so
# the data is preserved instead of silently dropped.
OP_RENAME = {
    "FileViewed":                "FileAccessed",
    "MeetingParticipantJoined":  "MeetingParticipantDetail",
    "ConnectedAIAppInteraction": "AIAppInteraction",
}

# Case-insensitive normalisation for every Operation we care about.
CANONICAL_OPS = KEEP_OPERATIONS | set(OP_RENAME.keys()) | {
    # Extras tolerated in raw exports (passed through normalisation,
    # then dropped by the KEEP filter — no behavioural change).
    "FilePreviewed", "MessagesListed", "ChatRetrieved",
    "MeetingDetail", "MeetingStarted", "MeetingEnded",
    "FileAccessedExtended", "CallParticipantDetail",
}
_OP_NORM = {op.lower(): op for op in CANONICAL_OPS}


def normalize_op(op: str) -> str:
    """Lowercase-match against CANONICAL_OPS, then apply OP_RENAME."""
    fixed = _OP_NORM.get((op or "").lower(), op)
    return OP_RENAME.get(fixed, fixed)


# ──────────────────────────────────────────────────────────────────────
# Agent classification (unchanged from build_rollup_with_agents.py)
# ──────────────────────────────────────────────────────────────────────

AGENT_APPHOSTS = {
    "Copilot Studio", "M365ChatApp", "BizChat", "ConnectedAIApp",
}
AGENT_OPERATIONS = {"AIAppInteraction"}


def parse_audit(raw):
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except (ValueError, TypeError):
        return {}


def extract_agent(audit, operation):
    ced = audit.get("CopilotEventData") or {}
    apphost = ced.get("AppHost") or ""
    agent_id = audit.get("AgentId") or ced.get("AgentId") or ""
    agent_name = audit.get("AgentName") or ""
    if not agent_name:
        plugins = ced.get("AISystemPlugin") or []
        if isinstance(plugins, list) and plugins:
            agent_name = (plugins[0] or {}).get("Name") or ""

    contexts = ced.get("Contexts") or []
    context_type = ""
    if isinstance(contexts, list) and contexts:
        context_type = (contexts[0] or {}).get("Type") or ""

    is_agent = bool(
        agent_id
        or operation in AGENT_OPERATIONS
        or apphost in AGENT_APPHOSTS
        or context_type == "agent"
    )
    return apphost, agent_id, agent_name, context_type, is_agent


def date_only(ts):
    if not ts:
        return ""
    try:
        return datetime.fromisoformat(ts.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        return ts[:10]


# ──────────────────────────────────────────────────────────────────────
# Core: stream one Purview CSV into the shared aggregator
# ──────────────────────────────────────────────────────────────────────

def consume_file(path: str, agg: dict, stats: dict) -> None:
    PROGRESS_EVERY = 100_000
    t0 = time.time()
    label = Path(path).name
    print(f"\n[{label}]  streaming...")

    rows_in = rows_kept = rows_agent = 0
    rows_renamed = rows_normalized = 0

    with open(path, "r", encoding="utf-8-sig", newline="") as fin:
        reader = csv.DictReader(fin)
        if "Operation" not in (reader.fieldnames or []):
            print(f"  !! '{label}' has no Operation column — skipping")
            return
        for row in reader:
            rows_in += 1
            if rows_in % PROGRESS_EVERY == 0:
                rate = rows_in / max(time.time() - t0, 0.001)
                print(f"  ... {rows_in:>12,} rows  ({rate:,.0f}/s)")

            raw_op = row.get("Operation") or ""
            op = normalize_op(raw_op)
            if op != raw_op:
                if raw_op in OP_RENAME:
                    rows_renamed += 1
                else:
                    rows_normalized += 1
            if op not in KEEP_OPERATIONS:
                continue

            audit = parse_audit(row.get("AuditData"))
            workload = audit.get("Workload") or ""
            source_ext = (
                audit.get("SourceFileExtension")
                or (audit.get("CopilotEventData") or {}).get("SourceFileExtension")
                or ""
            )
            apphost, agent_id, agent_name, ctx, is_agent = extract_agent(audit, op)

            ts = audit.get("CreationTime") or row.get("CreationDate") or ""
            day = date_only(ts)
            user = row.get("UserId") or ""

            key = (user, day, op, workload, source_ext, apphost,
                   agent_id, agent_name, ctx, is_agent)
            slot = agg[key]
            slot[0] += 1
            if not slot[1] or ts < slot[1]:
                slot[1] = ts
            if ts > slot[2]:
                slot[2] = ts

            rows_kept += 1
            if is_agent:
                rows_agent += 1

    elapsed = time.time() - t0
    print(f"  read   : {rows_in:,}    kept: {rows_kept:,} "
          f"({rows_kept / max(rows_in, 1) * 100:.1f}%)   agent: {rows_agent:,}   "
          f"renamed: {rows_renamed:,}   case-fixed: {rows_normalized:,}   "
          f"[{elapsed:.0f}s]")

    stats["rows_in"]         += rows_in
    stats["rows_kept"]       += rows_kept
    stats["rows_agent"]      += rows_agent
    stats["rows_renamed"]    += rows_renamed
    stats["rows_normalized"] += rows_normalized


# ──────────────────────────────────────────────────────────────────────
# Driver
# ──────────────────────────────────────────────────────────────────────

OUT_COLS = [
    "UserId", "CreationDate", "Operation", "Workload",
    "SourceFileExtension", "AppHost",
    "EventCount", "CreationTime", "MaxCreationTime",
    "AgentId", "AgentName", "ContextType", "IsAgentInteraction",
]


def write_rollup(agg: dict, output_path: str) -> None:
    with open(output_path, "w", encoding="utf-8", newline="") as fout:
        w = csv.writer(fout)
        w.writerow(OUT_COLS)
        for key, vals in agg.items():
            (user, day, op, workload, source_ext, apphost,
             agent_id, agent_name, ctx, is_agent) = key
            event_count, creation_time, max_time = vals
            w.writerow([
                user, day, op, workload, source_ext, apphost,
                event_count, creation_time, max_time,
                agent_id, agent_name, ctx,
                "true" if is_agent else "false",
            ])


def main(argv):
    if len(argv) < 3:
        print("Usage: python build_rollup_combined.py <output.csv> "
              "<input1.csv> [input2.csv ...]")
        sys.exit(1)

    output_path = argv[1]
    input_paths = argv[2:]

    for p in input_paths:
        if not Path(p).is_file():
            print(f"ERROR: input not found: {p}")
            sys.exit(2)

    agg = defaultdict(lambda: [0, "", ""])
    stats = defaultdict(int)
    t0 = time.time()

    for p in input_paths:
        consume_file(p, agg, stats)

    write_rollup(agg, output_path)

    elapsed = time.time() - t0
    print()
    print("=" * 66)
    print(f"Inputs           : {len(input_paths)}")
    print(f"Raw rows read    : {stats['rows_in']:,}")
    print(f"Rows kept        : {stats['rows_kept']:,} "
          f"({stats['rows_kept'] / max(stats['rows_in'], 1) * 100:.1f}%)")
    print(f"Rows renamed     : {stats['rows_renamed']:,}  "
          f"(FileViewed/MPJ/ConnectedAIApp -> canonical)")
    print(f"Rows case-fixed  : {stats['rows_normalized']:,}")
    print(f"Agent rows       : {stats['rows_agent']:,}")
    print(f"Rollup rows out  : {len(agg):,}")
    print(f"Output           : {output_path}")
    print(f"Total time       : {elapsed:.0f}s")


if __name__ == "__main__":
    main(sys.argv)
