"""
build_rollup_with_agents.py
────────────────────────────────────────────────────────────────────────────
Fixed Purview → Rollup pipeline that PRESERVES agent telemetry.

Output schema = the existing rollup schema (so the Power BI model's
column-fingerprint validator continues to pass) PLUS four agent columns
appended at the end:

   UserId, CreationDate, Operation, Workload, SourceFileExtension,
   AppHost, EventCount, CreationTime, MaxCreationTime,
   AgentId, AgentName, ContextType, IsAgentInteraction      ← NEW

Group-by key now includes AgentId / ContextType / IsAgentInteraction so
agent rows survive aggregation instead of being collapsed under
AgentId=null.

Usage:
   python build_rollup_with_agents.py <input_raw.csv> <output_rollup.csv>
"""
import csv
import json
import sys
import time
from collections import defaultdict
from datetime import datetime

csv.field_size_limit(10 * 1024 * 1024)   # AuditData JSON can be large

AGENT_APPHOSTS = {
    "Copilot Studio", "M365ChatApp", "BizChat", "ConnectedAIApp",
}
AGENT_OPERATIONS = {
    "ConnectedAIAppInteraction",
}

# Operations we emit. Adding ConnectedAIAppInteraction is one of the
# two core fixes (the previous pipeline filtered these rows out).
KEEP_OPERATIONS = {
    "CopilotInteraction",
    "ConnectedAIAppInteraction",
    "FileAccessed", "FileModified", "FileDownloaded", "FileUploaded",
    "FileViewed",
    "MessageSent", "MessageRead",
    "MailItemsAccessed", "Send", "MailboxLogin",
    "MeetingParticipantJoined", "TeamsSessionStarted", "ChatCreated",
}

# Canonical CamelCase forms for every Operation the PBIT or CLO dashboards
# read. Purview has been observed emitting some of these in lowercase
# ("send", "mailitemsaccessed", "mailboxlogin"), which silently breaks
# case-sensitive DAX IN {...} filters downstream. Normalizing at intake
# prevents the bug from recurring regardless of source casing.
CANONICAL_OPS = {
    "CopilotInteraction", "ConnectedAIAppInteraction",
    "MailItemsAccessed", "MailboxLogin", "Send",
    "FileAccessed", "FileViewed", "FilePreviewed", "FileModified",
    "FileDownloaded", "FileUploaded",
    "MessageSent", "MessageRead", "MessagesListed",
    "ChatRetrieved", "ChatCreated",
    "MeetingParticipantJoined", "MeetingStarted", "MeetingEnded",
    "MeetingParticipantDetail", "MeetingDetail",
    "TeamsSessionStarted",
}
_OP_NORM = {op.lower(): op for op in CANONICAL_OPS}


def normalize_op(op):
    return _OP_NORM.get((op or "").lower(), op)


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


def rollup(input_path, output_path):
    # value = [EventCount, min_CreationTime, MaxCreationTime]
    agg = defaultdict(lambda: [0, "", ""])

    n_in = n_kept = n_agent = 0
    n_with_agentid = 0
    n_connected = 0
    n_case_normalized = 0
    t0 = time.time()
    PROGRESS_EVERY = 100_000

    with open(input_path, "r", encoding="utf-8", newline="") as fin:
        reader = csv.DictReader(fin)
        for row in reader:
            n_in += 1
            if n_in % PROGRESS_EVERY == 0:
                rate = n_in / max(time.time() - t0, 0.001)
                print(f"  ... {n_in:>12,} rows  ({rate:,.0f} rows/sec) "
                      f"kept={n_kept:,} agent={n_agent:,}")

            raw_op = row.get("Operation") or ""
            op = normalize_op(raw_op)
            if op != raw_op:
                n_case_normalized += 1
            if op not in KEEP_OPERATIONS:
                continue

            audit = parse_audit(row.get("AuditData"))
            workload = audit.get("Workload") or ""
            source_ext = (
                audit.get("SourceFileExtension")
                or (audit.get("CopilotEventData") or {}).get("SourceFileExtension")
                or ""
            )

            apphost, agent_id, agent_name, context_type, is_agent = extract_agent(audit, op)

            ts = audit.get("CreationTime") or row.get("CreationDate") or ""
            day = date_only(ts)
            user = row.get("UserId") or ""

            # Agent fields MUST be in the group-by key; otherwise agent rows
            # get collapsed into AgentId="" buckets and the signal is lost.
            key = (user, day, op, workload, source_ext, apphost,
                   agent_id, agent_name, context_type, is_agent)

            slot = agg[key]
            slot[0] += 1
            if not slot[1] or ts < slot[1]:
                slot[1] = ts
            if ts > slot[2]:
                slot[2] = ts

            n_kept += 1
            if is_agent:
                n_agent += 1
            if agent_id:
                n_with_agentid += 1
            if op == "ConnectedAIAppInteraction":
                n_connected += 1

    out_cols = [
        "UserId", "CreationDate", "Operation", "Workload",
        "SourceFileExtension", "AppHost",
        "EventCount", "CreationTime", "MaxCreationTime",
        "AgentId", "AgentName", "ContextType", "IsAgentInteraction",
    ]
    with open(output_path, "w", encoding="utf-8", newline="") as fout:
        w = csv.writer(fout)
        w.writerow(out_cols)
        for key, vals in agg.items():
            (user, day, op, workload, source_ext, apphost,
             agent_id, agent_name, context_type, is_agent) = key
            event_count, creation_time, max_time = vals
            w.writerow([
                user, day, op, workload, source_ext, apphost,
                event_count, creation_time, max_time,
                agent_id, agent_name, context_type,
                "true" if is_agent else "false",
            ])

    elapsed = time.time() - t0
    print()
    print(f"Read    : {n_in:,} raw audit rows  ({elapsed:.0f}s)")
    print(f"Kept    : {n_kept:,} ({n_kept / max(n_in, 1) * 100:.1f}%)")
    print(f"Normaliz: {n_case_normalized:,} rows had Operation case-corrected")
    print(f"Agent   : {n_agent:,} agent rows  ({n_agent / max(n_kept, 1) * 100:.1f}% of kept)")
    print(f"           - with AgentId         : {n_with_agentid:,}")
    print(f"           - ConnectedAIAppInter. : {n_connected:,}")
    print(f"Rollup  : {len(agg):,} aggregated rows -> {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python build_rollup_with_agents.py <input.csv> <output.csv>")
        sys.exit(1)
    rollup(sys.argv[1], sys.argv[2])
