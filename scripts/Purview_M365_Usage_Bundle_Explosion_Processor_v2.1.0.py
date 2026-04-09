#!/usr/bin/env python3
"""
Purview M365 Usage Bundle Explosion Processor v2.1.0
=====================================================
Two-mode processor for Purview audit log CSV exports:

  ROLLUP MODE (default):  Aggregates exploded events into rolled-up rows keyed by
      (UserId, CreationDate, Operation, Workload, SourceFileExtension, AppHost)
      with EventCount, MIN(CreationTime), MAX(CreationTime).  Targets 80%+ row
      reduction for Power BI ingestion.  Streaming — no exploded rows held in memory.

      After the rollup CSV is written, a second pass streams through it to produce
      two additional analytics files (unless --no-userstats is specified):
        - UserStats:      One row per user with 27 columns of pre-computed metrics
                          (Copilot/M365 event counts, tier classifications, priority
                          scores, usage ranks, active-day counts, activity segments).
        - SessionCohort:  One row per (UserId, App) pair with a session-count bucket
                          (1-5, 6-10, 11-20, 21-40, 41-60, 61-80, 81+).

      These files allow Power Query to join pre-computed results instead of
      recalculating expensive DAX/M expressions, cutting dashboard load times.

  EVENT-LEVEL MODE (--mode event-level):  v1-compatible 153-column explosion output.
      Identical behavior to v1.0.0 for debugging and reconciliation.
      UserStats and SessionCohort files are NOT generated in this mode.

Requirements:
    Python 3.9+
    pip install orjson   (OPTIONAL - 5-10x faster JSON parsing; falls back to stdlib json)

Usage:
    python Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py --input <CSV>
        [--output-dir <DIR>] [--mode rollup|event-level]
        [--prompt-filter Prompt|Response|Both|Null]
        [--reconcile] [--no-userstats] [--quiet] [--version]

Output files (rollup mode — all share the same timestamp):
    <input_stem>_Rollup_<YYYYMMDD_HHMMSS>.csv          9 columns — aggregated events
    <input_stem>_UserStats_<YYYYMMDD_HHMMSS>.csv       27 columns — per-user metrics
    <input_stem>_SessionCohort_<YYYYMMDD_HHMMSS>.csv    3 columns — (UserId, App, Bucket)

Output file (event-level mode):
    <input_stem>_Exploded_<YYYYMMDD_HHMMSS>.csv       153 columns — one row per event

Arguments:
    --input, -i           Path to Purview audit log CSV (required).
    --output-dir, -o      Directory for output files (default: input file's directory).
    --mode, -m            Processing mode: rollup (default) or event-level.
    --reconcile           Run sample-based reconciliation after rollup processing.
    --prompt-filter       Filter Copilot messages: Prompt|Response|Both|Null.
    --no-userstats        Skip UserStats and SessionCohort generation (rollup only).
    --quiet, -q           Suppress progress output (only errors are printed).
    --version             Show version and exit.

Examples:
    # Default rollup (9-column output + UserStats + SessionCohort)
    python Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py -i Purview_Export.csv

    # Rollup with output in a different directory
    python Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py -i Purview_Export.csv --output-dir ./output

    # Rollup only — skip UserStats and SessionCohort generation
    python Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py -i Purview_Export.csv --no-userstats

    # v1-compatible event-level explosion (153-column output)
    python Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py -i Purview_Export.csv --mode event-level

    # Rollup with sample-based reconciliation check
    python Purview_M365_Usage_Bundle_Explosion_Processor_v2.1.0.py -i Purview_Export.csv --reconcile

Author:  Microsoft Copilot Growth ROI Advisory Team (copilot-roi-advisory-team-gh@microsoft.com)
Version: 2.1.0
"""

from __future__ import annotations

import argparse
import csv
import os
import random
import sys
import time
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

# ─── Fast JSON: prefer orjson, fall back to stdlib ───────────────────────────
try:
    import orjson

    def json_loads(s: str | bytes) -> Any:
        if isinstance(s, str):
            s = s.encode("utf-8")
        return orjson.loads(s)

    def json_dumps_compact(obj: Any) -> str:
        return orjson.dumps(obj, option=orjson.OPT_NON_STR_KEYS).decode("utf-8")

    _JSON_ENGINE = "orjson"
except ImportError:
    import json as _json

    def json_loads(s: str | bytes) -> Any:  # type: ignore[misc]
        if isinstance(s, bytes):
            s = s.decode("utf-8")
        return _json.loads(s)

    def json_dumps_compact(obj: Any) -> str:  # type: ignore[misc]
        return _json.dumps(obj, separators=(",", ":"), default=str)

    _JSON_ENGINE = "json (stdlib)"

# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════

SCRIPT_VERSION = "2.1.0"

EXPLOSION_PER_RECORD_ROW_CAP = 1000
STREAMING_CHUNK_SIZE = 5000

# Unified 153-column header matching Power BI M code schema exactly.
# Order matches the #"Changed Type" step in M365Usage.tmdl.
# AuditData is intentionally excluded — raw JSON is never written to output.
M365_UNIFIED_HEADER: list[str] = [
    "RecordId", "CreationDate", "RecordType", "Operation", "UserId",
    "AssociatedAdminUnits", "AssociatedAdminUnitsNames",
    "@odata.type", "CreationTime", "Id", "OrganizationId",
    "ResultStatus", "UserKey", "UserType", "Version", "Workload",
    "ClientIP", "ObjectId", "AzureActiveDirectoryEventType",
    "ActorContextId", "ActorIpAddress", "InterSystemsId", "IntraSystemId",
    "SupportTicketId", "TargetContextId", "ApplicationId",
    "DeviceProperties.OS", "DeviceProperties.BrowserType",
    "ErrorNumber",
    "SiteUrl", "SourceRelativeUrl", "SourceFileName", "SourceFileExtension",
    "ListId", "ListItemUniqueId", "WebId", "ApplicationDisplayName", "EventSource",
    "ItemType", "SiteSensitivityLabelId", "GeoLocation", "IsManagedDevice",
    "DeviceDisplayName", "ListBaseType", "ListServerTemplate",
    "AuthenticationType", "Site", "DoNotDistributeEvent", "HighPriorityMediaProcessing",
    "BrowserName", "BrowserVersion", "CorrelationId", "Platform", "UserAgent",
    "ActorInfoString", "AppId", "AuthType", "ClientAppId", "ClientIPAddress",
    "ClientInfoString", "ExternalAccess", "InternalLogonType", "LogonType",
    "LogonUserSid", "MailboxGuid", "MailboxOwnerSid", "MailboxOwnerUPN",
    "OrganizationName", "OriginatingServer", "SessionId",
    "TokenObjectId", "TokenTenantId", "TokenType", "SaveToSentItems",
    "OperationCount", "FileSizeBytes",
    "MeetingId", "MeetingType", "EventSignature", "EventData",
    "Permission", "SensitivityLabelId", "SharingLinkScope",
    "TargetUserOrGroupType", "TargetUserOrGroupName",
    "MeetingURL", "ChatId", "MessageId", "MessageSizeInBytes", "MessageType",
    "FormId", "FormName", "VideoId", "VideoName", "ChannelId", "ViewDuration",
    "ClientRegion", "CopilotLogVersion", "TargetId",
    "TeamName", "TeamGuid", "ResponseId", "IsAnonymous", "DeviceType",
    "ChannelName", "ChannelGuid", "ChannelType", "AppName", "EnvironmentName",
    "PlanId", "PlanName", "TaskId", "TaskName", "PercentComplete",
    "CrossMailboxOperation",
    "RecordTypeNum", "ResultStatus_Audit",
    "ModelId", "ModelProvider", "ModelFamily",
    "TokensTotal", "TokensInput", "TokensOutput", "DurationMs", "OutcomeStatus",
    "ConversationId", "TurnNumber", "RetryCount", "ClientVersion", "ClientPlatform",
    "AgentId", "AgentName", "AgentVersion", "AgentCategory", "ApplicationName",
    "AppHost", "ThreadId",
    "Context_Id", "Context_Type",
    "Message_Id", "Message_isPrompt",
    "AccessedResource_Action", "AccessedResource_PolicyDetails", "AccessedResource_SiteUrl",
    "AISystemPlugin_Id", "AISystemPlugin_Name",
    "ModelTransparencyDetails_ModelName", "MessageIds",
    "AccessedResource_Name", "AccessedResource_SensitivityLabel",
    "AccessedResource_ResourceType", "SensitivityLabel", "Context_Item",
]

# Rollup output header (9 columns)
ROLLUP_HEADER: list[str] = [
    "UserId", "CreationDate", "Operation", "Workload",
    "SourceFileExtension", "AppHost",
    "EventCount", "CreationTime", "MaxCreationTime",
]

# Reconciliation sample size
RECONCILE_SAMPLE_SIZE = 10_000

# ── UserStats classification sets (match Power Query logic exactly) ──────────
WORD_EXTS: set[str] = {"docx", "doc", "dotx"}
EXCEL_EXTS: set[str] = {"xlsx", "xls", "xlsm", "csv"}
PPT_EXTS: set[str] = {"pptx", "ppt", "ppsx"}
OFFICE_EXTS: set[str] = WORD_EXTS | EXCEL_EXTS | PPT_EXTS

FILE_OPS: set[str] = {"FileViewed", "FileModified", "FileDownloaded", "FileUploaded"}
OUTLOOK_OPS: set[str] = {"Send", "MailItemsAccessed", "MailboxLogin"}           # active-DAY counting
OUTLOOK_ACT_OPS: set[str] = {"Send", "MailItemsAccessed", "MailboxLogin"}      # event COUNT
TEAMS_OPS: set[str] = {"MessageSent", "MessageRead", "MeetingParticipantJoined"}

USERSTATS_HEADER: list[str] = [
    "UserId",
    "CopilotEC", "M365EC", "ExCopEC", "ExM365EC",
    "IsCopilotUser", "CopilotTierColumn", "M365TierColumn",
    "PriorityScatterColumn", "ExcelPriority",
    "CopilotUsageRankColumn", "M365UsageRankColumn",
    "TeamsActiveDays", "OutlookActiveDays", "WordActiveDays",
    "ExcelActiveDays", "PowerPointActiveDays",
    "TeamsActivityCount", "OutlookActivityCount", "OfficeFilesActivityCount",
    "TeamsActivitySegment", "OutlookActivitySegment", "WordActivitySegment",
    "ExcelActivitySegment", "PowerPointActivitySegment",
    "OfficeFilesActivitySegment", "OverallM365ActivitySegment",
]

SESSIONCOHORT_HEADER: list[str] = ["UserId", "AppColumn", "SessionCohort"]

# Date formats accepted for CreationDate normalization (broadest to narrowest)
_CREATION_DATE_FORMATS: tuple[str, ...] = (
    "%Y-%m-%dT%H:%M:%S.%fZ",
    "%Y-%m-%dT%H:%M:%SZ",
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S",
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y %H:%M:%S",
    "%Y-%m-%d",
    "%m/%d/%Y",
)

# GroupKey type: (user_id_lower, creation_date_normalized, operation, workload, sfe_lower, app_host)
GroupKey = tuple[str, str, str, str, str, str]


class RollupAccum:
    """Lightweight accumulator for one rollup group — avoids dataclass import overhead."""
    __slots__ = ("event_count", "min_creation_time", "max_creation_time", "original_user_id")

    def __init__(self, event_count: int, min_ct: str, max_ct: str, original_uid: str) -> None:
        self.event_count = event_count
        self.min_creation_time = min_ct
        self.max_creation_time = max_ct
        self.original_user_id = original_uid  # first-seen casing for output


def normalize_creation_date(raw: str) -> str:
    """Parse any Purview date format → 'YYYY-MM-DDT00:00:00.000Z' (midnight UTC)."""
    if not raw or not isinstance(raw, str):
        return ""
    raw = raw.strip()
    if not raw:
        return ""
    for fmt in _CREATION_DATE_FORMATS:
        try:
            dt = datetime.strptime(raw, fmt)
            return dt.strftime("%Y-%m-%d") + "T00:00:00.000Z"
        except ValueError:
            continue
    # Fallback: try extracting date portion from ISO-like string
    if len(raw) >= 10 and raw[4:5] == "-":
        return raw[:10] + "T00:00:00.000Z"
    return raw  # unparseable — pass through


def _norm_key_str(val: Any) -> str:
    """Normalize a string value for use as a rollup key: strip whitespace, empty if None."""
    if val is None:
        return ""
    if not isinstance(val, str):
        val = str(val)
    val = val.strip()
    if val.lower() in ("null", "none"):
        return ""
    return val


# ═════════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════

def safe_get(obj: Any, key: str) -> Any:
    """Safely retrieve a property from a dict-like object."""
    if obj is None:
        return None
    if isinstance(obj, dict):
        return obj.get(key)
    return getattr(obj, key, None)


def select_first_non_null(values: list[Any]) -> Any:
    """Return the first non-None, non-empty-string value."""
    for v in values:
        if v is not None and v != "":
            return v
    return None


def to_num(val: Any) -> float | None:
    """Convert to number, return None on failure."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        val = val.strip()
        if not val:
            return None
        try:
            return float(val)
        except (ValueError, TypeError):
            return None
    return None


def format_date_purview(val: Any) -> str:
    """Format a date value to ISO 8601 UTC string."""
    if val is None:
        return ""
    if isinstance(val, str):
        val = val.strip()
        if not val:
            return ""
        # Try common Purview date formats
        for fmt in (
            "%Y-%m-%dT%H:%M:%S.%fZ",
            "%Y-%m-%dT%H:%M:%SZ",
            "%Y-%m-%dT%H:%M:%S.%f",
            "%Y-%m-%dT%H:%M:%S",
            "%m/%d/%Y %I:%M:%S %p",
            "%m/%d/%Y %H:%M:%S",
        ):
            try:
                dt = datetime.strptime(val, fmt)
                return dt.replace(tzinfo=timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"
            except ValueError:
                continue
        return val  # Return as-is if no format matches
    return str(val)


def to_json_if_object(val: Any) -> str:
    """Serialize non-scalars to compact JSON, pass scalars through as strings."""
    if val is None:
        return ""
    if isinstance(val, (str, int, float, bool)):
        return str(val)
    try:
        return json_dumps_compact(val)
    except Exception:
        return ""


def bool_tf(val: Any) -> str:
    """Convert bool-like value to 'TRUE'/'FALSE' string."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, str):
        low = val.strip().lower()
        if low in ("true", "1", "yes"):
            return "TRUE"
        if low in ("false", "0", "no"):
            return "FALSE"
    return str(val)


def get_array_fast(obj: Any, key: str) -> list:
    """Extract an array property, always returning a list."""
    if obj is None:
        return []
    val = safe_get(obj, key)
    if val is None:
        return []
    if isinstance(val, list):
        return val
    if isinstance(val, (str, int, float, bool)):
        return []
    try:
        return list(val)
    except (TypeError, ValueError):
        return []


# ═════════════════════════════════════════════════════════════════════════════
# AGENT CATEGORIZATION
# ═════════════════════════════════════════════════════════════════════════════

def categorize_agent(agent_id: Any) -> str:
    """Categorize agent based on AgentId pattern."""
    if not agent_id or not isinstance(agent_id, str):
        return ""
    if agent_id.startswith("CopilotStudio.Declarative."):
        return "Declarative Agent"
    if agent_id.startswith("CopilotStudio.CustomEngine."):
        return "Custom Engine Agent"
    if agent_id.startswith("P_"):
        return "Declarative Agent (Purview)"
    return "Other Agent"


# ═════════════════════════════════════════════════════════════════════════════
# USERSTATS CLASSIFICATION & COMPUTATION HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def is_copilot(op: str, wl: str) -> bool:
    """True if the row represents a Copilot event."""
    return wl == "Copilot" or op == "CopilotInteraction"


def is_excel_file_op(ext: str, op: str) -> bool:
    """True if the row is a file operation on an Excel-family extension."""
    return (ext or "").lower() in EXCEL_EXTS and op in FILE_OPS


def app_column(ext: str, op: str, wl: str) -> str:
    """Classify a row into an application column for session cohort grouping."""
    e = (ext or "").lower()
    if e in WORD_EXTS and op in FILE_OPS:
        return "Word"
    if e in EXCEL_EXTS and op in FILE_OPS:
        return "Excel"
    if e in PPT_EXTS and op in FILE_OPS:
        return "PowerPoint"
    if wl == "Exchange" and op in OUTLOOK_ACT_OPS:
        return "Outlook"
    if wl == "MicrosoftTeams" and op in TEAMS_OPS:
        return "Teams"
    if wl == "Copilot" or op == "CopilotInteraction":
        return "Copilot"
    return "M365 All Apps"


def percentile_inc(values: list[float], p: float) -> float:
    """Inclusive linear interpolation — matches PQ List.Percentile and numpy 'linear'."""
    if not values:
        return 0.0
    sorted_vals = sorted(values)
    n = len(sorted_vals)
    idx = p * (n - 1)
    lo, hi = int(idx), min(int(idx) + 1, n - 1)
    return sorted_vals[lo] + (idx - lo) * (sorted_vals[hi] - sorted_vals[lo])


def tier_fn(cnt: float, p90: float, p75: float, p50: float, zero_is_bottom: bool) -> str:
    """Assign a percentile-tier label."""
    if zero_is_bottom and cnt == 0:
        return "Bottom 50%"
    if cnt >= p90:
        return "Top 10%"
    if cnt >= p75:
        return "10-25%"
    if cnt >= p50:
        return "25-50%"
    return "Bottom 50%"


def priority_fn(m365_tier: str, cop_tier: str) -> str:
    """Map (M365 tier, Copilot tier) pair to a priority label."""
    top2 = {"Top 10%", "10-25%"}
    if m365_tier in top2 and cop_tier in top2:
        return "Promoter"
    if m365_tier == "Top 10%" and cop_tier == "25-50%":
        return "High"
    if m365_tier == "Top 10%" and cop_tier == "Bottom 50%":
        return "Critical"
    if m365_tier == "10-25%" and cop_tier == "25-50%":
        return "Medium"
    if m365_tier == "10-25%" and cop_tier == "Bottom 50%":
        return "High"
    if m365_tier in {"25-50%", "Bottom 50%"} and cop_tier in top2:
        return "Promoter"
    if m365_tier == "25-50%" and cop_tier == "25-50%":
        return "Medium"
    if m365_tier == "25-50%" and cop_tier == "Bottom 50%":
        return "Medium"
    return "Low"


def seg_fn(days: int) -> str:
    """Map active-day count to an activity-segment label."""
    if days <= 5:
        return "1. 1-5 Days (Infrequent)"
    if days <= 10:
        return "2. 6-10 Days (Moderate)"
    if days <= 19:
        return "3. 11-19 Days (Frequent)"
    return "4. 20+ Days (Daily)"


def compute_ranks(values_by_uid: dict[str, float]) -> dict[str, int]:
    """
    0-based ascending rank via stable sort.  Ties get different sequential
    indices (matches PQ Table.Sort + Table.AddIndexColumn behaviour).
    """
    sorted_uids = sorted(values_by_uid.items(), key=lambda x: x[1])
    return {uid: i for i, (uid, _) in enumerate(sorted_uids)}


# ═════════════════════════════════════════════════════════════════════════════
# ROLLUP KEY EXTRACTION (lightweight — no row dicts built)
# ═════════════════════════════════════════════════════════════════════════════

def _compute_copilot_event_count(
    ced: dict,
    operation: str,
    prompt_filter: str | None,
) -> int:
    """
    Compute the number of exploded rows a Copilot record would produce,
    using the same array-length logic as v1 (explode_copilot_record),
    but WITHOUT materializing any row dicts.
    Returns 0 if prompt_filter eliminates all messages (record skipped).
    """
    messages = get_array_fast(ced, "Messages")
    contexts = get_array_fast(ced, "Contexts")
    resources = get_array_fast(ced, "AccessedResources")
    plugins_raw = get_array_fast(ced, "AISystemPlugin")
    model_det_raw = get_array_fast(ced, "ModelTransparencyDetails")
    sensitivity_labels = get_array_fast(ced, "SensitivityLabels")

    # Prompt filtering (same logic as v1 lines 571-581)
    if prompt_filter:
        pf_lower = prompt_filter.lower()
        if pf_lower == "null":
            messages = [m for m in messages if safe_get(m, "isPrompt") is None]
        elif pf_lower == "both":
            messages = [m for m in messages if safe_get(m, "isPrompt") is not None]
        elif pf_lower == "prompt":
            messages = [m for m in messages if safe_get(m, "isPrompt") is True]
        elif pf_lower == "response":
            messages = [m for m in messages if safe_get(m, "isPrompt") is False]
        if not messages:
            return 0  # record filtered out entirely

    # Context items max for CopilotInteraction
    context_items_max: int = 0
    if operation == "CopilotInteraction" and contexts:
        for ctx in contexts:
            if ctx:
                items = get_array_fast(ctx, "Items")
                if items and len(items) > context_items_max:
                    context_items_max = len(items)

    if prompt_filter:
        row_count = max(1, len(messages))
    else:
        array_counts = [
            1, len(messages), len(contexts), len(resources),
            len(sensitivity_labels), len(plugins_raw), len(model_det_raw),
        ]
        if context_items_max > 0:
            array_counts.append(context_items_max)
        row_count = max(array_counts)

    return min(max(row_count, 1), EXPLOSION_PER_RECORD_ROW_CAP)


def _extract_rollup_keys(
    record: dict,
    audit_data: dict,
    ced: dict | None,
    prompt_filter: str | None = None,
) -> tuple[GroupKey, int, str, str] | None:
    """
    Extract rollup group key + event count + creation_time + original UserId.

    Returns None if the record is filtered out (e.g. prompt_filter eliminates all messages).
    Returns:
        (group_key, event_count, creation_time_iso, original_user_id)
    where group_key uses lowercased UserId for case-insensitive grouping.
    """
    # UserId: original casing preserved for output; lowered for grouping key
    raw_uid = _norm_key_str(safe_get(audit_data, "UserId") or record.get("UserId", ""))
    uid_lower = raw_uid.lower()

    # CreationDate: from CSV, normalized to midnight
    creation_date = normalize_creation_date(record.get("CreationDate", ""))

    # Operation: from audit_data → CSV fallback, preserve case
    operation = _norm_key_str(
        safe_get(audit_data, "Operation") or record.get("Operation", "") or record.get("Operations", "")
    )

    # Workload: from audit_data, preserve case
    workload = _norm_key_str(safe_get(audit_data, "Workload"))

    # SourceFileExtension: lowercased for DAX LOWER() compatibility
    sfe = _norm_key_str(safe_get(audit_data, "SourceFileExtension")).lower()

    # AppHost: from CED for Copilot, empty otherwise, preserve case
    if ced:
        app_host = _norm_key_str(
            safe_get(ced, "AppHost") or safe_get(audit_data, "AppHost")
        )
    else:
        app_host = ""

    # CreationTime: from audit_data, ISO formatted for lexicographic MIN/MAX
    creation_time = format_date_purview(safe_get(audit_data, "CreationTime"))

    # Event count
    if ced:
        event_count = _compute_copilot_event_count(ced, operation, prompt_filter)
        if event_count == 0:
            return None  # filtered out
    else:
        event_count = 1  # Non-Copilot: always 1:1

    group_key: GroupKey = (uid_lower, creation_date, operation, workload, sfe, app_host)
    return group_key, event_count, creation_time, raw_uid


# ═════════════════════════════════════════════════════════════════════════════
# PATH A: NON-COPILOT M365 EXTRACTION
# ═════════════════════════════════════════════════════════════════════════════


def _get_nv_prop(nv_list: Any, prop_name: str) -> Any:
    """Extract a value from a Name/Value pair list by Name — matches M code's GetNVProp."""
    if not nv_list or not isinstance(nv_list, list):
        return None
    for item in nv_list:
        if isinstance(item, dict) and safe_get(item, "Name") == prop_name:
            return safe_get(item, "Value")
    return None


def _build_unified_row(record: dict, audit_data: dict) -> dict:
    """
    Build a complete row dict with all 153 M code columns populated.
    Extracts fields from the CSV record and AuditData JSON.
    DeviceProperties uses NV-pivot for .OS and .BrowserType only (matches M code).
    RecordTypeNum and ResultStatus_Audit are computed aliases.
    """
    # CSV-level fields
    record_id = (
        record.get("RecordId")
        or record.get("Identity")
        or record.get("Id")
        or safe_get(audit_data, "Id")
        or ""
    )
    # Read from singular first, fall back to plural for backwards-compatible input
    op_val = safe_get(audit_data, "Operation") or record.get("Operation") or record.get("Operations", "")
    uid_val = safe_get(audit_data, "UserId") or record.get("UserId") or record.get("UserIds", "")
    record_type = record.get("RecordType", "")
    result_status = safe_get(audit_data, "ResultStatus") or ""

    # CreationTime formatting
    creation_time_raw = safe_get(audit_data, "CreationTime")
    creation_time = format_date_purview(creation_time_raw) if creation_time_raw else ""

    # DeviceProperties NV-pivot (matches M code's GetNVProp — only .OS and .BrowserType)
    dev_props = safe_get(audit_data, "DeviceProperties")
    dp_os = _get_nv_prop(dev_props, "OS") or ""
    dp_browser = _get_nv_prop(dev_props, "BrowserType") or ""

    # Computed alias: RecordTypeNum = int(RecordType)
    try:
        record_type_num = int(record_type) if record_type else ""
    except (ValueError, TypeError):
        record_type_num = ""

    # ApplicationId with fallback chain
    app_id_resolved = select_first_non_null([
        safe_get(audit_data, "ApplicationId"),
        safe_get(audit_data, "AppId"),
        safe_get(audit_data, "ClientAppId"),
    ]) or ""

    # AgentCategory is computed
    agent_id_val = safe_get(audit_data, "AgentId") or ""
    agent_category = categorize_agent(agent_id_val) if agent_id_val else ""

    row = {
        "RecordId": record_id,
        "CreationDate": record.get("CreationDate", ""),
        "RecordType": record_type,
        "Operation": op_val,
        "UserId": uid_val,
        "AssociatedAdminUnits": record.get("AssociatedAdminUnits", "") or safe_get(audit_data, "AssociatedAdminUnits") or "",
        "AssociatedAdminUnitsNames": record.get("AssociatedAdminUnitsNames", "") or safe_get(audit_data, "AssociatedAdminUnitsNames") or "",
        "@odata.type": safe_get(audit_data, "@odata.type") or "",
        "CreationTime": creation_time,
        "Id": safe_get(audit_data, "Id") or "",
        "OrganizationId": safe_get(audit_data, "OrganizationId") or "",
        "ResultStatus": result_status,
        "UserKey": safe_get(audit_data, "UserKey") or "",
        "UserType": safe_get(audit_data, "UserType") or "",
        "Version": safe_get(audit_data, "Version") or "",
        "Workload": safe_get(audit_data, "Workload") or "",
        "ClientIP": safe_get(audit_data, "ClientIP") or "",
        "ObjectId": safe_get(audit_data, "ObjectId") or "",
        "AzureActiveDirectoryEventType": safe_get(audit_data, "AzureActiveDirectoryEventType") or "",
        "ActorContextId": safe_get(audit_data, "ActorContextId") or "",
        "ActorIpAddress": safe_get(audit_data, "ActorIpAddress") or "",
        "InterSystemsId": safe_get(audit_data, "InterSystemsId") or "",
        "IntraSystemId": safe_get(audit_data, "IntraSystemId") or "",
        "SupportTicketId": safe_get(audit_data, "SupportTicketId") or "",
        "TargetContextId": safe_get(audit_data, "TargetContextId") or "",
        "ApplicationId": app_id_resolved,
        "DeviceProperties.OS": dp_os,
        "DeviceProperties.BrowserType": dp_browser,
        "ErrorNumber": safe_get(audit_data, "ErrorNumber") or "",
        "SiteUrl": safe_get(audit_data, "SiteUrl") or "",
        "SourceRelativeUrl": safe_get(audit_data, "SourceRelativeUrl") or "",
        "SourceFileName": safe_get(audit_data, "SourceFileName") or "",
        "SourceFileExtension": safe_get(audit_data, "SourceFileExtension") or "",
        "ListId": safe_get(audit_data, "ListId") or "",
        "ListItemUniqueId": safe_get(audit_data, "ListItemUniqueId") or "",
        "WebId": safe_get(audit_data, "WebId") or "",
        "ApplicationDisplayName": safe_get(audit_data, "ApplicationDisplayName") or "",
        "EventSource": safe_get(audit_data, "EventSource") or "",
        "ItemType": safe_get(audit_data, "ItemType") or "",
        "SiteSensitivityLabelId": safe_get(audit_data, "SiteSensitivityLabelId") or "",
        "GeoLocation": safe_get(audit_data, "GeoLocation") or "",
        "IsManagedDevice": safe_get(audit_data, "IsManagedDevice") or "",
        "DeviceDisplayName": safe_get(audit_data, "DeviceDisplayName") or "",
        "ListBaseType": safe_get(audit_data, "ListBaseType") or "",
        "ListServerTemplate": safe_get(audit_data, "ListServerTemplate") or "",
        "AuthenticationType": safe_get(audit_data, "AuthenticationType") or "",
        "Site": safe_get(audit_data, "Site") or "",
        "DoNotDistributeEvent": safe_get(audit_data, "DoNotDistributeEvent") or "",
        "HighPriorityMediaProcessing": safe_get(audit_data, "HighPriorityMediaProcessing") or "",
        "BrowserName": safe_get(audit_data, "BrowserName") or "",
        "BrowserVersion": safe_get(audit_data, "BrowserVersion") or "",
        "CorrelationId": safe_get(audit_data, "CorrelationId") or "",
        "Platform": safe_get(audit_data, "Platform") or "",
        "UserAgent": safe_get(audit_data, "UserAgent") or "",
        "ActorInfoString": safe_get(audit_data, "ActorInfoString") or "",
        "AppId": safe_get(audit_data, "AppId") or "",
        "AuthType": safe_get(audit_data, "AuthType") or "",
        "ClientAppId": safe_get(audit_data, "ClientAppId") or "",
        "ClientIPAddress": safe_get(audit_data, "ClientIPAddress") or "",
        "ClientInfoString": safe_get(audit_data, "ClientInfoString") or "",
        "ExternalAccess": safe_get(audit_data, "ExternalAccess") or "",
        "InternalLogonType": safe_get(audit_data, "InternalLogonType") or "",
        "LogonType": safe_get(audit_data, "LogonType") or "",
        "LogonUserSid": safe_get(audit_data, "LogonUserSid") or "",
        "MailboxGuid": safe_get(audit_data, "MailboxGuid") or "",
        "MailboxOwnerSid": safe_get(audit_data, "MailboxOwnerSid") or "",
        "MailboxOwnerUPN": safe_get(audit_data, "MailboxOwnerUPN") or "",
        "OrganizationName": safe_get(audit_data, "OrganizationName") or "",
        "OriginatingServer": safe_get(audit_data, "OriginatingServer") or "",
        "SessionId": safe_get(audit_data, "SessionId") or "",
        "TokenObjectId": safe_get(audit_data, "TokenObjectId") or "",
        "TokenTenantId": safe_get(audit_data, "TokenTenantId") or "",
        "TokenType": safe_get(audit_data, "TokenType") or "",
        "SaveToSentItems": safe_get(audit_data, "SaveToSentItems") or "",
        "OperationCount": safe_get(audit_data, "OperationCount") or "",
        "FileSizeBytes": safe_get(audit_data, "FileSizeBytes") or "",
        # Teams / Meetings / Chat
        "MeetingId": safe_get(audit_data, "MeetingId") or "",
        "MeetingType": safe_get(audit_data, "MeetingType") or "",
        "EventSignature": safe_get(audit_data, "EventSignature") or "",
        "EventData": safe_get(audit_data, "EventData") or "",
        "Permission": safe_get(audit_data, "Permission") or "",
        "SensitivityLabelId": safe_get(audit_data, "SensitivityLabelId") or "",
        "SharingLinkScope": safe_get(audit_data, "SharingLinkScope") or "",
        "TargetUserOrGroupType": safe_get(audit_data, "TargetUserOrGroupType") or "",
        "TargetUserOrGroupName": safe_get(audit_data, "TargetUserOrGroupName") or "",
        "MeetingURL": safe_get(audit_data, "MeetingURL") or "",
        "ChatId": safe_get(audit_data, "ChatId") or "",
        "MessageId": safe_get(audit_data, "MessageId") or "",
        "MessageSizeInBytes": safe_get(audit_data, "MessageSizeInBytes") or "",
        "MessageType": safe_get(audit_data, "MessageType") or "",
        # Forms
        "FormId": safe_get(audit_data, "FormId") or "",
        "FormName": safe_get(audit_data, "FormName") or "",
        # Video / Stream
        "VideoId": safe_get(audit_data, "VideoId") or "",
        "VideoName": safe_get(audit_data, "VideoName") or "",
        "ChannelId": safe_get(audit_data, "ChannelId") or "",
        "ViewDuration": safe_get(audit_data, "ViewDuration") or "",
        "ClientRegion": safe_get(audit_data, "ClientRegion") or "",
        "CopilotLogVersion": safe_get(audit_data, "CopilotLogVersion") or "",
        "TargetId": safe_get(audit_data, "TargetId") or "",
        # Teams details
        "TeamName": safe_get(audit_data, "TeamName") or "",
        "TeamGuid": safe_get(audit_data, "TeamGuid") or "",
        "ResponseId": safe_get(audit_data, "ResponseId") or "",
        "IsAnonymous": safe_get(audit_data, "IsAnonymous") or "",
        "DeviceType": safe_get(audit_data, "DeviceType") or "",
        "ChannelName": safe_get(audit_data, "ChannelName") or "",
        "ChannelGuid": safe_get(audit_data, "ChannelGuid") or "",
        "ChannelType": safe_get(audit_data, "ChannelType") or "",
        "AppName": safe_get(audit_data, "AppName") or "",
        "EnvironmentName": safe_get(audit_data, "EnvironmentName") or "",
        # Planner
        "PlanId": safe_get(audit_data, "PlanId") or "",
        "PlanName": safe_get(audit_data, "PlanName") or "",
        "TaskId": safe_get(audit_data, "TaskId") or "",
        "TaskName": safe_get(audit_data, "TaskName") or "",
        "PercentComplete": safe_get(audit_data, "PercentComplete") or "",
        "CrossMailboxOperation": safe_get(audit_data, "CrossMailboxOperation") or "",
        # Computed aliases
        "RecordTypeNum": record_type_num,
        "ResultStatus_Audit": result_status,
        # Copilot model/token fields (populated from CED for Copilot records; from root for non-Copilot)
        "ModelId": safe_get(audit_data, "ModelId") or "",
        "ModelProvider": safe_get(audit_data, "ModelProvider") or "",
        "ModelFamily": safe_get(audit_data, "ModelFamily") or "",
        "TokensTotal": safe_get(audit_data, "TokensTotal") or "",
        "TokensInput": safe_get(audit_data, "TokensInput") or "",
        "TokensOutput": safe_get(audit_data, "TokensOutput") or "",
        "DurationMs": safe_get(audit_data, "DurationMs") or "",
        "OutcomeStatus": safe_get(audit_data, "OutcomeStatus") or "",
        "ConversationId": safe_get(audit_data, "ConversationId") or "",
        "TurnNumber": safe_get(audit_data, "TurnNumber") or "",
        "RetryCount": safe_get(audit_data, "RetryCount") or "",
        "ClientVersion": safe_get(audit_data, "ClientVersion") or "",
        "ClientPlatform": safe_get(audit_data, "ClientPlatform") or "",
        "AgentId": agent_id_val,
        "AgentName": safe_get(audit_data, "AgentName") or "",
        "AgentVersion": safe_get(audit_data, "AgentVersion") or "",
        "AgentCategory": agent_category,
        "ApplicationName": safe_get(audit_data, "ApplicationName") or "",
        "SensitivityLabel": safe_get(audit_data, "SensitivityLabel") or "",
        # CED sub-fields — empty for non-Copilot records, populated by Copilot path
        "AppHost": "",
        "ThreadId": "",
        "Context_Id": "",
        "Context_Type": "",
        "Message_Id": "",
        "Message_isPrompt": "",
        "AccessedResource_Action": "",
        "AccessedResource_PolicyDetails": "",
        "AccessedResource_SiteUrl": "",
        "AISystemPlugin_Id": "",
        "AISystemPlugin_Name": "",
        "ModelTransparencyDetails_ModelName": "",
        "MessageIds": "",
        "AccessedResource_Name": "",
        "AccessedResource_SensitivityLabel": "",
        "AccessedResource_ResourceType": "",
        "Context_Item": "",
    }
    return row


def explode_m365_record(record: dict, audit_data: dict) -> list[dict]:
    """
    Extract a non-Copilot M365 record (Path A).
    Produces exactly 1 row per record with all 153 M code columns.
    No array explosion — M code does not explode non-Copilot arrays.
    """
    return [_build_unified_row(record, audit_data)]


# ═════════════════════════════════════════════════════════════════════════════
# PATH B: COPILOT EXPLOSION
# ═════════════════════════════════════════════════════════════════════════════

def explode_copilot_record(
    record: dict,
    audit_data: dict,
    ced: dict,
    prompt_filter: str | None = None,
) -> list[dict]:
    """
    Explode a Copilot record (Path B).
    Starts from the unified 153-column base row, then overrides CED-specific fields.
    Extracts Messages, Contexts, AccessedResources, AISystemPlugin,
    ModelTransparencyDetails, SensitivityLabels and builds N parallel-indexed rows.
    """
    # Extract array fields from CopilotEventData
    messages = get_array_fast(ced, "Messages")
    contexts = get_array_fast(ced, "Contexts")
    resources = get_array_fast(ced, "AccessedResources")
    plugins_raw = get_array_fast(ced, "AISystemPlugin")
    model_det_raw = get_array_fast(ced, "ModelTransparencyDetails")
    message_ids = get_array_fast(ced, "MessageIds")
    sensitivity_labels = get_array_fast(ced, "SensitivityLabels")

    # Prompt filtering
    if prompt_filter:
        filtered: list = []
        pf_lower = prompt_filter.lower()
        if pf_lower == "null":
            filtered = [m for m in messages if safe_get(m, "isPrompt") is None]
        elif pf_lower == "both":
            filtered = [m for m in messages if safe_get(m, "isPrompt") is not None]
        elif pf_lower == "prompt":
            filtered = [m for m in messages if safe_get(m, "isPrompt") is True]
        elif pf_lower == "response":
            filtered = [m for m in messages if safe_get(m, "isPrompt") is False]
        messages = filtered
        if not messages:
            return []

    # Detect activity type for 2-level explosion
    activity_type = safe_get(audit_data, "Operation") or ""

    # Context items max for CopilotInteraction
    context_items_max: int = 0
    if activity_type == "CopilotInteraction" and contexts:
        for ctx in contexts:
            if ctx:
                items = get_array_fast(ctx, "Items")
                if items and len(items) > context_items_max:
                    context_items_max = len(items)

    # Calculate row count
    if prompt_filter:
        row_count = max(1, len(messages))
    else:
        array_counts = [
            1, len(messages), len(contexts), len(resources),
            len(sensitivity_labels), len(plugins_raw), len(model_det_raw),
        ]
        if context_items_max > 0:
            array_counts.append(context_items_max)
        row_count = max(array_counts)

    row_count = min(row_count, EXPLOSION_PER_RECORD_ROW_CAP)
    if row_count < 1:
        row_count = 1

    # ── Build unified base row with all 153 M code columns ───────────────
    base = _build_unified_row(record, audit_data)

    # ── Override CED-specific scalar fields with deep CED extraction ─────
    # AppHost: prefer CED → audit_data → Workload
    base["AppHost"] = select_first_non_null([
        safe_get(ced, "AppHost"),
        safe_get(audit_data, "AppHost"),
        safe_get(audit_data, "Workload"),
    ]) or ""

    base["ThreadId"] = safe_get(ced, "ThreadId") or ""

    # AgentVersion: prefer audit_data → CED fallbacks
    base["AgentVersion"] = select_first_non_null([
        safe_get(audit_data, "AgentVersion"),
        safe_get(ced, "AgentVersion"),
        safe_get(ced, "Version"),
    ]) or ""

    # ApplicationName: prefer audit_data → CED fallbacks
    base["ApplicationName"] = select_first_non_null([
        safe_get(audit_data, "ApplicationName"),
        safe_get(ced, "HostAppName"),
        safe_get(ced, "ClientAppName"),
    ]) or ""

    # Model fields from CED with fallbacks
    base["ModelId"] = select_first_non_null([
        safe_get(ced, "ModelId"), safe_get(ced, "ModelID"), safe_get(audit_data, "ModelId"),
    ]) or ""
    base["ModelProvider"] = select_first_non_null([
        safe_get(ced, "ModelProvider"), safe_get(ced, "Provider"), safe_get(ced, "ModelVendor"),
    ]) or ""
    base["ModelFamily"] = select_first_non_null([
        safe_get(ced, "ModelFamily"), safe_get(ced, "ModelType"),
    ]) or ""

    # Token usage from CED
    usage_node = select_first_non_null([
        safe_get(ced, "Usage"), safe_get(ced, "TokenUsage"),
        safe_get(ced, "Tokens"), safe_get(audit_data, "Usage"),
    ])
    tokens_total: Any = None
    tokens_input: Any = None
    tokens_output: Any = None
    if usage_node and isinstance(usage_node, dict):
        tokens_total = to_num(select_first_non_null([
            safe_get(usage_node, "Total"), safe_get(usage_node, "TotalTokens"),
            safe_get(usage_node, "TokensTotal"),
        ]))
        tokens_input = to_num(select_first_non_null([
            safe_get(usage_node, "Input"), safe_get(usage_node, "Prompt"),
            safe_get(usage_node, "InputTokens"), safe_get(usage_node, "TokensInput"),
        ]))
        tokens_output = to_num(select_first_non_null([
            safe_get(usage_node, "Output"), safe_get(usage_node, "Completion"),
            safe_get(usage_node, "OutputTokens"), safe_get(usage_node, "TokensOutput"),
        ]))
    if not tokens_total and (tokens_input or tokens_output):
        try:
            tokens_total = (tokens_input or 0) + (tokens_output or 0)
        except Exception:
            pass
    base["TokensTotal"] = tokens_total if tokens_total is not None else ""
    base["TokensInput"] = tokens_input if tokens_input is not None else ""
    base["TokensOutput"] = tokens_output if tokens_output is not None else ""

    # Duration, outcome, conversation from CED
    duration_ms = to_num(select_first_non_null([
        safe_get(ced, "DurationMs"), safe_get(ced, "ElapsedMs"),
        safe_get(ced, "ProcessingTimeMs"), safe_get(ced, "LatencyMs"),
    ]))
    base["DurationMs"] = duration_ms if duration_ms is not None else ""

    outcome_status: Any = select_first_non_null([
        safe_get(ced, "OutcomeStatus"), safe_get(ced, "Outcome"),
        safe_get(ced, "Result"), safe_get(ced, "Status"),
    ])
    if isinstance(outcome_status, bool):
        outcome_status = "Success" if outcome_status else "Failure"
    base["OutcomeStatus"] = outcome_status or ""

    base["ConversationId"] = select_first_non_null([
        safe_get(ced, "ConversationId"), safe_get(ced, "ConversationID"),
        safe_get(ced, "SessionId"),
    ]) or ""

    turn_number = to_num(select_first_non_null([
        safe_get(ced, "TurnNumber"), safe_get(ced, "TurnIndex"),
        safe_get(ced, "MessageIndex"),
    ]))
    base["TurnNumber"] = turn_number if turn_number is not None else ""

    retry_count = to_num(select_first_non_null([
        safe_get(ced, "RetryCount"), safe_get(ced, "Retries"),
    ]))
    base["RetryCount"] = retry_count if retry_count is not None else ""

    base["ClientVersion"] = select_first_non_null([
        safe_get(ced, "ClientVersion"), safe_get(ced, "Version"), safe_get(ced, "Build"),
    ]) or ""
    base["ClientPlatform"] = select_first_non_null([
        safe_get(ced, "ClientPlatform"), safe_get(ced, "Platform"), safe_get(ced, "OS"),
    ]) or ""

    # MessageIds: semicolon-joined (matches M code's Text.Combine)
    base["MessageIds"] = ";".join(str(m) for m in message_ids) if message_ids else ""

    # ── Build rows with indexed array access ─────────────────────────────
    rows: list[dict] = []
    for i in range(row_count):
        row = dict(base)  # shallow copy of all 153 columns

        # Indexed array access — Contexts
        if i < len(contexts) and contexts[i]:
            row["Context_Id"] = safe_get(contexts[i], "Id") or ""
            row["Context_Type"] = safe_get(contexts[i], "Type") or ""
        else:
            row["Context_Id"] = ""
            row["Context_Type"] = ""

        # Messages
        if i < len(messages):
            msg = messages[i]
            if isinstance(msg, dict):
                row["Message_Id"] = safe_get(msg, "Id") or ""
                row["Message_isPrompt"] = bool_tf(safe_get(msg, "isPrompt"))
            else:
                row["Message_Id"] = str(msg) if msg is not None else ""
                row["Message_isPrompt"] = ""
        else:
            row["Message_Id"] = ""
            row["Message_isPrompt"] = ""

        # AccessedResources
        if i < len(resources) and resources[i]:
            res = resources[i]
            row["AccessedResource_Action"] = safe_get(res, "Action") or ""
            row["AccessedResource_PolicyDetails"] = to_json_if_object(safe_get(res, "PolicyDetails"))
            row["AccessedResource_SiteUrl"] = safe_get(res, "SiteUrl") or ""
            row["AccessedResource_Name"] = safe_get(res, "Name") or ""
            row["AccessedResource_SensitivityLabel"] = safe_get(res, "SensitivityLabel") or ""
            row["AccessedResource_ResourceType"] = safe_get(res, "ResourceType") or ""
        else:
            row["AccessedResource_Action"] = ""
            row["AccessedResource_PolicyDetails"] = ""
            row["AccessedResource_SiteUrl"] = ""
            row["AccessedResource_Name"] = ""
            row["AccessedResource_SensitivityLabel"] = ""
            row["AccessedResource_ResourceType"] = ""

        # AISystemPlugin
        if i < len(plugins_raw) and plugins_raw[i]:
            row["AISystemPlugin_Id"] = safe_get(plugins_raw[i], "Id") or ""
            row["AISystemPlugin_Name"] = safe_get(plugins_raw[i], "Name") or ""
        else:
            row["AISystemPlugin_Id"] = ""
            row["AISystemPlugin_Name"] = ""

        # ModelTransparencyDetails
        if i < len(model_det_raw) and model_det_raw[i]:
            row["ModelTransparencyDetails_ModelName"] = safe_get(model_det_raw[i], "ModelName") or ""
        else:
            row["ModelTransparencyDetails_ModelName"] = ""

        # SensitivityLabel (from CED SensitivityLabels array)
        if i < len(sensitivity_labels):
            row["SensitivityLabel"] = str(sensitivity_labels[i]) if sensitivity_labels[i] is not None else ""

        # Context_Item — full mode: one item per row across all contexts
        if activity_type == "CopilotInteraction":
            found_item = None
            for ctx in contexts:
                if ctx:
                    items = get_array_fast(ctx, "Items")
                    if items and i < len(items):
                        found_item = items[i]
                        break
            row["Context_Item"] = to_json_if_object(found_item) if found_item else ""
        else:
            row["Context_Item"] = ""

        rows.append(row)

    return rows


# ═════════════════════════════════════════════════════════════════════════════
# ROUTER: Dispatch to Path A or Path B
# ═════════════════════════════════════════════════════════════════════════════

def explode_record(
    record: dict,
    prompt_filter: str | None = None,
) -> list[dict]:
    """
    Parse AuditData and route to appropriate explosion path.
    Returns list of flattened row dicts, or empty list on error.
    """
    audit_data_raw = record.get("AuditData", "")
    if not audit_data_raw or not isinstance(audit_data_raw, str) or not audit_data_raw.strip():
        return []

    try:
        audit_data = json_loads(audit_data_raw)
    except Exception:
        return []

    if not isinstance(audit_data, dict):
        return []

    ced = safe_get(audit_data, "CopilotEventData")
    if ced and isinstance(ced, dict):
        return explode_copilot_record(record, audit_data, ced, prompt_filter=prompt_filter)
    else:
        return explode_m365_record(record, audit_data)


# ═════════════════════════════════════════════════════════════════════════════
# HEADER — Fixed schema, no dynamic discovery needed
# ═════════════════════════════════════════════════════════════════════════════
# Output columns are exactly M365_UNIFIED_HEADER (153 columns in M code order).
# No schema discovery pass is needed because both Path A and Path B produce
# row dicts that contain exactly these keys.


# ═════════════════════════════════════════════════════════════════════════════
# CHUNK PROCESSOR (unit of parallel work)
# ═════════════════════════════════════════════════════════════════════════════

def _process_chunk(args: tuple) -> tuple[list[dict], int, int]:
    """
    Process a chunk of CSV rows → exploded row dicts.
    Returns (exploded_rows, input_count, error_count).
    """
    chunk, prompt_filter = args
    results: list[dict] = []
    errors = 0

    for record in chunk:
        try:
            rows = explode_record(record, prompt_filter=prompt_filter)
            results.extend(rows)
        except Exception:
            errors += 1

    return results, len(chunk), errors


# ═════════════════════════════════════════════════════════════════════════════
# MAIN EXPLOSION ORCHESTRATOR
# ═════════════════════════════════════════════════════════════════════════════

def run_explosion(
    input_csv: str,
    output_csv: str,
    prompt_filter: str | None = None,
    workers: int = 0,
    chunk_size: int = STREAMING_CHUNK_SIZE,
    quiet: bool = False,
) -> dict[str, Any]:
    """
    Main entry point: reads input CSV, explodes all records, writes output CSV.
    Uses multiprocessing for large files, single-process for small ones.

    Returns a stats dict with counts and timing.
    """
    if not os.path.isfile(input_csv):
        print(f"ERROR: Input file not found: {input_csv}", file=sys.stderr)
        sys.exit(1)

    if workers <= 0:
        workers = min(os.cpu_count() or 1, 8)

    t_start = time.perf_counter()
    stats = {
        "input_records": 0,
        "output_rows": 0,
        "errors": 0,
        "chunks_processed": 0,
    }

    if not quiet:
        print(f"Purview M365 Usage Bundle Explosion Processor v{SCRIPT_VERSION}")
        print(f"  JSON engine:    {_JSON_ENGINE}")
        print(f"  Input:          {input_csv}")
        print(f"  Output:         {output_csv}")
        print(f"  Prompt filter:  {prompt_filter or 'None'}")
        print(f"  Workers:        {workers}")
        print(f"  Chunk size:     {chunk_size}")
        print()

    # ── Phase 1: Fixed schema ─────────────────────────────────────────────
    final_header = list(M365_UNIFIED_HEADER)  # 153 columns in M code order
    if not quiet:
        print(f"Phase 1: Using fixed {len(final_header)}-column M code schema")

    # ── Phase 2: Process chunks ──────────────────────────────────────────
    if not quiet:
        print("Phase 2: Processing records...")

    # Accumulate all exploded rows (we need dynamic columns before writing header)
    all_rows: list[dict] = []

    # Read CSV in chunks
    chunks: list[list[dict]] = []
    current_chunk: list[dict] = []

    with open(input_csv, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            current_chunk.append(row)
            if len(current_chunk) >= chunk_size:
                chunks.append(current_chunk)
                current_chunk = []
        if current_chunk:
            chunks.append(current_chunk)

    total_input = sum(len(c) for c in chunks)
    stats["input_records"] = total_input

    if not quiet:
        print(f"  Loaded {total_input:,} input records in {len(chunks)} chunk(s)")

    # Determine whether to use multiprocessing
    use_parallel = workers > 1 and len(chunks) > 1

    if use_parallel:
        chunk_args = [(chunk, prompt_filter) for chunk in chunks]
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(_process_chunk, arg): idx for idx, arg in enumerate(chunk_args)}
            for future in as_completed(futures):
                try:
                    exploded, _in_count, err_count = future.result()
                    all_rows.extend(exploded)
                    stats["errors"] += err_count
                    stats["chunks_processed"] += 1
                    if not quiet and stats["chunks_processed"] % 5 == 0:
                        print(f"    Chunks completed: {stats['chunks_processed']}/{len(chunks)}")
                except Exception as exc:
                    stats["errors"] += 1
                    if not quiet:
                        print(f"    Chunk failed: {exc}", file=sys.stderr)
    else:
        for chunk in chunks:
            exploded, _in_count, err_count = _process_chunk((chunk, prompt_filter))
            all_rows.extend(exploded)
            stats["errors"] += err_count
            stats["chunks_processed"] += 1
            if not quiet and stats["chunks_processed"] % 5 == 0:
                print(f"    Chunks completed: {stats['chunks_processed']}/{len(chunks)}")

    stats["output_rows"] = len(all_rows)

    # ── Phase 3: Write output CSV with fixed header ──────────────────────
    if not quiet:
        print("Phase 3: Writing output CSV...")

    # Write output using fixed M code-ordered header
    os.makedirs(os.path.dirname(os.path.abspath(output_csv)), exist_ok=True)
    with open(output_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=final_header, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in all_rows:
            writer.writerow(row)

    t_elapsed = time.perf_counter() - t_start

    # ── Summary ──────────────────────────────────────────────────────────
    if not quiet:
        print()
        print("=== EXPLOSION SUMMARY ===")
        print(f"  Input records:  {stats['input_records']:,}")
        print(f"  Output rows:    {stats['output_rows']:,}")
        if stats["output_rows"] > stats["input_records"] and stats["input_records"] > 0:
            ratio = round(stats["output_rows"] / stats["input_records"], 2)
            extra = stats["output_rows"] - stats["input_records"]
            print(f"  Expansion:      {ratio}x ({extra:,} additional rows from array explosion)")
        elif stats["output_rows"] == stats["input_records"]:
            print("  Expansion:      1:1 (no arrays exploded)")
        elif stats["input_records"] > 0:
            filtered = stats["input_records"] - stats["output_rows"]
            print(f"  Reduction:      {filtered:,} records filtered out")
        if stats["errors"] > 0:
            print(f"  Errors:         {stats['errors']:,} record(s) failed to process")
        print(f"  Columns:        {len(final_header):,}")
        print(f"  Elapsed:        {t_elapsed:.2f}s")
        if stats["output_rows"] > 0 and t_elapsed > 0:
            print(f"  Throughput:     {stats['output_rows'] / t_elapsed:,.0f} rows/sec")
        print(f"  Output file:    {output_csv}")
        print()

    return stats


# ═════════════════════════════════════════════════════════════════════════════
# ROLLUP ORCHESTRATOR (streaming — no exploded rows in memory)
# ═════════════════════════════════════════════════════════════════════════════

def run_rollup(
    input_csv: str,
    output_csv: str,
    prompt_filter: str | None = None,
    quiet: bool = False,
) -> dict[str, Any]:
    """
    Streaming rollup: read CSV row-by-row → parse AuditData → extract 6 keys +
    CreationTime → accumulate into dict[GroupKey, RollupAccum] → write 9-column CSV.

    No exploded row dicts are ever stored in memory.
    """
    if not os.path.isfile(input_csv):
        print(f"ERROR: Input file not found: {input_csv}", file=sys.stderr)
        sys.exit(1)

    t_start = time.perf_counter()
    rollup: dict[GroupKey, RollupAccum] = {}
    stats: dict[str, Any] = {
        "input_records": 0,
        "virtual_exploded_event_count": 0,
        "output_rows": 0,
        "parse_errors": 0,
    }

    if not quiet:
        print(f"Purview M365 Usage Bundle Explosion Processor v{SCRIPT_VERSION} [ROLLUP MODE]")
        print(f"  JSON engine:    {_JSON_ENGINE}")
        print(f"  Input:          {input_csv}")
        print(f"  Output:         {output_csv}")
        print(f"  Prompt filter:  {prompt_filter or 'None'}")
        print()
        print("Processing records (streaming rollup)...")

    # ── Streaming read + accumulate ──────────────────────────────────────
    with open(input_csv, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for record in reader:
            stats["input_records"] += 1

            # Progress indicator
            if not quiet and stats["input_records"] % 500_000 == 0:
                print(f"  {stats['input_records']:>12,} records processed, "
                      f"{len(rollup):,} groups...")

            # Parse AuditData JSON
            audit_data_raw = record.get("AuditData", "")
            if not audit_data_raw or not isinstance(audit_data_raw, str) or not audit_data_raw.strip():
                stats["parse_errors"] += 1
                continue
            try:
                audit_data = json_loads(audit_data_raw)
            except Exception:
                stats["parse_errors"] += 1
                continue
            if not isinstance(audit_data, dict):
                stats["parse_errors"] += 1
                continue

            ced = safe_get(audit_data, "CopilotEventData")
            if ced and not isinstance(ced, dict):
                ced = None

            # Extract rollup keys (lightweight — no row dict built)
            result = _extract_rollup_keys(record, audit_data, ced, prompt_filter)
            if result is None:
                continue  # filtered out by prompt_filter

            group_key, event_count, creation_time, original_uid = result
            stats["virtual_exploded_event_count"] += event_count

            # Accumulate into rollup dict
            if group_key in rollup:
                acc = rollup[group_key]
                acc.event_count += event_count
                if creation_time:
                    if not acc.min_creation_time or creation_time < acc.min_creation_time:
                        acc.min_creation_time = creation_time
                    if not acc.max_creation_time or creation_time > acc.max_creation_time:
                        acc.max_creation_time = creation_time
            else:
                rollup[group_key] = RollupAccum(
                    event_count=event_count,
                    min_ct=creation_time,
                    max_ct=creation_time,
                    original_uid=original_uid,
                )

    # ── Write rollup output CSV ──────────────────────────────────────────
    stats["output_rows"] = len(rollup)

    if not quiet:
        print(f"  {stats['input_records']:>12,} records processed (done)")
        print(f"  Writing {stats['output_rows']:,} rollup rows...")

    os.makedirs(os.path.dirname(os.path.abspath(output_csv)), exist_ok=True)
    with open(output_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(ROLLUP_HEADER)
        for (uid_lower, cdate, op, wl, sfe, ah), acc in rollup.items():
            writer.writerow([
                acc.original_user_id,  # output original casing, NOT lowered key
                cdate,
                op,
                wl,
                sfe,
                ah,
                acc.event_count,
                acc.min_creation_time,   # CreationTime = MIN
                acc.max_creation_time,   # MaxCreationTime = MAX
            ])

    t_elapsed = time.perf_counter() - t_start

    # ── Summary report ───────────────────────────────────────────────────
    if not quiet:
        pct = 0.0
        if stats["virtual_exploded_event_count"] > 0:
            pct = (1 - stats["output_rows"] / stats["virtual_exploded_event_count"]) * 100
        print()
        print("=== ROLLUP SUMMARY ===")
        print(f"  Input records:              {stats['input_records']:>14,}")
        print(f"  Virtual exploded events:    {stats['virtual_exploded_event_count']:>14,}")
        print(f"  Rollup output rows:         {stats['output_rows']:>14,}")
        print(f"  Row reduction:              {pct:>13.1f}%"
              f"  ({stats['virtual_exploded_event_count']:,} -> {stats['output_rows']:,})")
        if stats["parse_errors"] > 0:
            print(f"  Parse errors:               {stats['parse_errors']:>14,}")
        print(f"  Columns:                    {len(ROLLUP_HEADER):>14}")
        print(f"  Elapsed:                    {t_elapsed:>13.2f}s")
        if stats["input_records"] > 0 and t_elapsed > 0:
            print(f"  Throughput:                 {stats['input_records'] / t_elapsed:>12,.0f} input records/sec")
        print(f"  Output file:                {output_csv}")
        print()

    return stats


# ═════════════════════════════════════════════════════════════════════════════
# RECONCILIATION (sample-based validation of rollup correctness)
# ═════════════════════════════════════════════════════════════════════════════

def run_reconcile(
    input_csv: str,
    prompt_filter: str | None = None,
    sample_size: int = RECONCILE_SAMPLE_SIZE,
    quiet: bool = False,
) -> bool:
    """
    Sample-based reconciliation: read a sample of records, run both rollup-key
    extraction and full event-level explosion, compare total and filtered counts.

    Returns True if all checks pass, False otherwise.
    """
    if not os.path.isfile(input_csv):
        print(f"ERROR: Input file not found: {input_csv}", file=sys.stderr)
        return False

    if not quiet:
        print(f"\n=== RECONCILIATION CHECK (sample {sample_size:,} records) ===\n")

    # ── Read sample ──────────────────────────────────────────────────────
    all_records: list[dict] = []
    with open(input_csv, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for record in reader:
            all_records.append(record)

    if len(all_records) > sample_size:
        sample = random.sample(all_records, sample_size)
    else:
        sample = all_records
        sample_size = len(sample)

    if not quiet:
        print(f"  Total input records: {len(all_records):,}")
        print(f"  Sample size:         {sample_size:,}")

    # ── Run event-level explosion on sample ──────────────────────────────
    event_rows: list[dict] = []
    event_errors = 0
    for record in sample:
        try:
            rows = explode_record(record, prompt_filter=prompt_filter)
            event_rows.extend(rows)
        except Exception:
            event_errors += 1

    # ── Run rollup-key extraction on same sample ─────────────────────────
    rollup_sample: dict[GroupKey, RollupAccum] = {}
    rollup_errors = 0
    for record in sample:
        audit_data_raw = record.get("AuditData", "")
        if not audit_data_raw or not isinstance(audit_data_raw, str) or not audit_data_raw.strip():
            rollup_errors += 1
            continue
        try:
            audit_data = json_loads(audit_data_raw)
        except Exception:
            rollup_errors += 1
            continue
        if not isinstance(audit_data, dict):
            rollup_errors += 1
            continue

        ced = safe_get(audit_data, "CopilotEventData")
        if ced and not isinstance(ced, dict):
            ced = None

        result = _extract_rollup_keys(record, audit_data, ced, prompt_filter)
        if result is None:
            continue
        group_key, event_count, creation_time, original_uid = result

        if group_key in rollup_sample:
            acc = rollup_sample[group_key]
            acc.event_count += event_count
            if creation_time:
                if not acc.min_creation_time or creation_time < acc.min_creation_time:
                    acc.min_creation_time = creation_time
                if not acc.max_creation_time or creation_time > acc.max_creation_time:
                    acc.max_creation_time = creation_time
        else:
            rollup_sample[group_key] = RollupAccum(
                event_count=event_count,
                min_ct=creation_time,
                max_ct=creation_time,
                original_uid=original_uid,
            )

    # ── Compare totals ───────────────────────────────────────────────────
    rollup_total = sum(acc.event_count for acc in rollup_sample.values())
    event_total = len(event_rows)
    all_pass = True

    def _check(label: str, rollup_val: Any, event_val: Any) -> bool:
        nonlocal all_pass
        match = rollup_val == event_val
        symbol = "PASS" if match else "FAIL"
        if not quiet:
            print(f"  {label}")
            print(f"    Rollup: {rollup_val}   Event-level: {event_val}   [{symbol}]")
        if not match:
            all_pass = False
        return match

    _check("Total event count (SUM(EventCount) vs COUNTROWS)",
           rollup_total, event_total)

    # ── Filter-specific checks ───────────────────────────────────────────
    # Helper: count event-level rows matching a filter
    def _ev_count(**filters: str | set) -> int:
        count = 0
        for row in event_rows:
            match = True
            for col, val in filters.items():
                row_val = row.get(col, "")
                if isinstance(val, set):
                    if row_val.lower() not in val:
                        match = False
                        break
                else:
                    if row_val != val:
                        match = False
                        break
            if match:
                count += 1
        return count

    # Helper: sum EventCount from rollup for matching groups
    def _ru_count(**filters: str | set) -> int:
        total = 0
        for (uid_l, cdate, op, wl, sfe, ah), acc in rollup_sample.items():
            match = True
            key_map = {"Operation": op, "Workload": wl,
                       "SourceFileExtension": sfe, "AppHost": ah}
            for col, val in filters.items():
                key_val = key_map.get(col, "")
                if isinstance(val, set):
                    if key_val.lower() not in val:
                        match = False
                        break
                else:
                    if key_val != val:
                        match = False
                        break
            if match:
                total += acc.event_count
        return total

    # Check 1: Teams MessageSent
    _check("Teams MessageSent (Workload=MicrosoftTeams, Operation=MessageSent)",
           _ru_count(Workload="MicrosoftTeams", Operation="MessageSent"),
           _ev_count(Workload="MicrosoftTeams", Operation="MessageSent"))

    # Check 2: Exchange Send
    _check("Exchange Send (Workload=Exchange, Operation=Send)",
           _ru_count(Workload="Exchange", Operation="Send"),
           _ev_count(Workload="Exchange", Operation="Send"))

    # Check 3: CopilotInteraction + AppHost=Teams
    _check("Copilot Teams (Operation=CopilotInteraction, AppHost=Teams)",
           _ru_count(Operation="CopilotInteraction", AppHost="Teams"),
           _ev_count(Operation="CopilotInteraction", AppHost="Teams"))

    # Check 4: Excel FileViewed
    _check("Excel FileViewed (Operation=FileViewed, SourceFileExtension in xlsx/xls/xlsm/csv)",
           _ru_count(Operation="FileViewed", SourceFileExtension={"xlsx", "xls", "xlsm", "csv"}),
           _ev_count(Operation="FileViewed", SourceFileExtension={"xlsx", "xls", "xlsm", "csv"}))

    # ── Temporal checks ──────────────────────────────────────────────────
    rollup_min_ct = min((acc.min_creation_time for acc in rollup_sample.values() if acc.min_creation_time), default="")
    rollup_max_ct = max((acc.max_creation_time for acc in rollup_sample.values() if acc.max_creation_time), default="")
    event_times = [r.get("CreationTime", "") for r in event_rows if r.get("CreationTime")]
    event_min_ct = min(event_times) if event_times else ""
    event_max_ct = max(event_times) if event_times else ""

    _check("MIN(CreationTime)", rollup_min_ct, event_min_ct)
    _check("MAX(CreationTime)", rollup_max_ct, event_max_ct)

    if not quiet:
        print()
        reduction_pct = 0.0
        if event_total > 0:
            reduction_pct = (1 - len(rollup_sample) / event_total) * 100
        print(f"  Event-level rows: {event_total:,}  →  Rollup groups: {len(rollup_sample):,}"
              f"  ({reduction_pct:.1f}% reduction)")
        print(f"  Overall: {'ALL CHECKS PASSED' if all_pass else 'SOME CHECKS FAILED'}")
        print()

    return all_pass


# ═════════════════════════════════════════════════════════════════════════════
# USERSTATS & SESSION COHORT WRITER
# ═════════════════════════════════════════════════════════════════════════════

def write_userstats_files(
    aggregated_csv_path: str | Path,
    userstats_csv_path: str | Path,
    session_csv_path: str | Path,
    quiet: bool,
) -> tuple[int, int]:
    """
    Read the just-written aggregated rollup CSV and produce two additional files:
      *_UserStats.csv     — one row per unique UserId with pre-computed metrics
      *_SessionCohort.csv — one row per (UserId, AppColumn) with session cohort label

    Returns (user_count, session_cohort_row_count).
    """
    agg_path = Path(aggregated_csv_path)
    if not agg_path.is_file():
        if not quiet:
            print(f"[UserStats] WARNING: Aggregated CSV not found: {agg_path} — skipping.",
                  file=sys.stderr)
        return 0, 0

    userstats_path = Path(userstats_csv_path)
    session_path = Path(session_csv_path)

    t_start = time.perf_counter()

    # ── Per-user accumulators ────────────────────────────────────────────
    uid_original: dict[str, str] = {}          # uid_lower → first-seen casing
    cop_ec: dict[str, int] = defaultdict(int)
    m365_ec: dict[str, int] = defaultdict(int)
    ex_cop_ec: dict[str, int] = defaultdict(int)
    ex_m365_ec: dict[str, int] = defaultdict(int)

    t_days: dict[str, set[str]] = defaultdict(set)
    o_days: dict[str, set[str]] = defaultdict(set)
    w_days: dict[str, set[str]] = defaultdict(set)
    x_days: dict[str, set[str]] = defaultdict(set)
    p_days: dict[str, set[str]] = defaultdict(set)

    t_ec: dict[str, int] = defaultdict(int)
    o_ec: dict[str, int] = defaultdict(int)
    off_ec: dict[str, int] = defaultdict(int)

    session_ops: dict[tuple[str, str], set[str]] = defaultdict(set)

    # ── Stream through aggregated CSV ────────────────────────────────────
    row_count = 0
    with open(agg_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            row_count += 1

            user_id = row.get("UserId", "")
            uid_lower = user_id.lower()
            if uid_lower not in uid_original:
                uid_original[uid_lower] = user_id

            date_key = row.get("CreationDate", "")[:10]   # YYYY-MM-DD
            op = row.get("Operation", "")
            wl = row.get("Workload", "")
            ext = (row.get("SourceFileExtension", "") or "").lower()
            app_host = (row.get("AppHost", "") or "").lower()

            try:
                event_count = int(row.get("EventCount", "1") or "1")
            except (ValueError, TypeError):
                event_count = 1

            copilot = is_copilot(op, wl)
            excel_file = is_excel_file_op(ext, op)

            # Core event counts
            if copilot:
                cop_ec[uid_lower] += event_count
            else:
                m365_ec[uid_lower] += event_count

            # ExCopEC: Copilot interactions in Excel (via AppHost); ExM365EC: Excel file ops by non-Copilot
            if copilot and app_host == "excel":
                ex_cop_ec[uid_lower] += 1        # row count, not EventCount
            if excel_file and not copilot:
                ex_m365_ec[uid_lower] += 1       # row count, not EventCount

            # Active days (distinct CreationDate values)
            if wl == "MicrosoftTeams" and op in TEAMS_OPS:
                t_days[uid_lower].add(date_key)
            if wl == "Exchange" and op in OUTLOOK_OPS:
                o_days[uid_lower].add(date_key)
            if ext in WORD_EXTS and op in FILE_OPS:
                w_days[uid_lower].add(date_key)
            if ext in EXCEL_EXTS and op in FILE_OPS:
                x_days[uid_lower].add(date_key)
            if ext in PPT_EXTS and op in FILE_OPS:
                p_days[uid_lower].add(date_key)

            # Activity event counts
            if wl == "MicrosoftTeams" and op in TEAMS_OPS:
                t_ec[uid_lower] += event_count
            if wl == "Exchange" and op in OUTLOOK_ACT_OPS:
                o_ec[uid_lower] += event_count
            if ext in OFFICE_EXTS and op in FILE_OPS:
                off_ec[uid_lower] += event_count

            # Session cohort: distinct active dates per (user, app)
            app = app_column(ext, op, wl)
            if app != "M365 All Apps":
                session_ops[(uid_lower, app)].add(date_key)

    if row_count == 0:
        if not quiet:
            print("[UserStats] Aggregated CSV has 0 rows — skipping.")
        return 0, 0

    # ── Percentile thresholds ────────────────────────────────────────────
    all_uids = sorted(uid_original.keys())
    total_users = len(all_uids)

    cop_vals = [cop_ec.get(u, 0) for u in all_uids]
    m365_vals = [m365_ec.get(u, 0) for u in all_uids]
    ex_cop_vals = [ex_cop_ec.get(u, 0) for u in all_uids]
    ex_m365_vals = [ex_m365_ec.get(u, 0) for u in all_uids]

    cop_p90 = percentile_inc(cop_vals, 0.90)
    cop_p75 = percentile_inc(cop_vals, 0.75)
    cop_p50 = percentile_inc(cop_vals, 0.50)

    m365_p90 = percentile_inc(m365_vals, 0.90)
    m365_p75 = percentile_inc(m365_vals, 0.75)
    m365_p50 = percentile_inc(m365_vals, 0.50)

    exc_p90 = percentile_inc(ex_cop_vals, 0.90)
    exc_p75 = percentile_inc(ex_cop_vals, 0.75)
    exc_p50 = percentile_inc(ex_cop_vals, 0.50)

    exm_p90 = percentile_inc(ex_m365_vals, 0.90)
    exm_p75 = percentile_inc(ex_m365_vals, 0.75)
    exm_p50 = percentile_inc(ex_m365_vals, 0.50)

    # ── Ranks ────────────────────────────────────────────────────────────
    # Copilot rank: computed only among Copilot users so the range maps to [0, 1]
    # within that group. Non-Copilot users are hardcoded to 0.0 downstream.
    copilot_uids = [u for u in all_uids if cop_ec.get(u, 0) > 0]
    copilot_user_count = len(copilot_uids)
    cop_rank = compute_ranks({u: cop_ec.get(u, 0) for u in copilot_uids})
    m365_rank = compute_ranks({u: m365_ec.get(u, 0) for u in all_uids})

    # ── Write *_UserStats.csv ────────────────────────────────────────────
    with open(userstats_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(USERSTATS_HEADER)

        for uid in all_uids:
            c_ec = cop_ec.get(uid, 0)
            m_ec = m365_ec.get(uid, 0)
            exc_ec = ex_cop_ec.get(uid, 0)
            exm_ec = ex_m365_ec.get(uid, 0)

            is_cop_user = "Copilot User" if c_ec > 0 else "Non-Copilot User"
            cop_tier = tier_fn(c_ec, cop_p90, cop_p75, cop_p50, zero_is_bottom=True)
            m365_tier = tier_fn(m_ec, m365_p90, m365_p75, m365_p50, zero_is_bottom=False)
            priority = priority_fn(m365_tier, cop_tier)

            ex_m365_tier = tier_fn(exm_ec, exm_p90, exm_p75, exm_p50, zero_is_bottom=False)
            ex_cop_tier = tier_fn(exc_ec, exc_p90, exc_p75, exc_p50, zero_is_bottom=False)
            excel_pri = priority_fn(ex_m365_tier, ex_cop_tier)

            cop_rank_val = 0.0 if c_ec == 0 else cop_rank[uid] / max(copilot_user_count, 1)
            m365_rank_val = m365_rank[uid] / total_users

            td = len(t_days.get(uid, set()))
            od = len(o_days.get(uid, set()))
            wd = len(w_days.get(uid, set()))
            xd = len(x_days.get(uid, set()))
            pd_ = len(p_days.get(uid, set()))

            t_act = t_ec.get(uid, 0)
            o_act = o_ec.get(uid, 0)
            off_act = off_ec.get(uid, 0)

            t_seg = "0. No Usage" if td == 0 else seg_fn(td)
            o_seg = "0. No Usage" if od == 0 else seg_fn(od)
            w_seg = "0. No Usage" if wd == 0 else seg_fn(wd)
            x_seg = "0. No Usage" if xd == 0 else seg_fn(xd)
            p_seg = "0. No Usage" if pd_ == 0 else seg_fn(pd_)

            office_days = wd + xd + pd_
            off_seg = "0. No Usage" if office_days == 0 else seg_fn(office_days)

            overall_days = len(
                t_days.get(uid, set()) | o_days.get(uid, set()) |
                w_days.get(uid, set()) | x_days.get(uid, set()) |
                p_days.get(uid, set())
            )
            overall_seg = "0. No Usage" if overall_days == 0 else seg_fn(overall_days)

            writer.writerow([
                uid_original[uid],
                c_ec, m_ec, exc_ec, exm_ec,
                is_cop_user, cop_tier, m365_tier,
                priority, excel_pri,
                f"{cop_rank_val:.6f}", f"{m365_rank_val:.6f}",
                td, od, wd, xd, pd_,
                t_act, o_act, off_act,
                t_seg, o_seg, w_seg, x_seg, p_seg,
                off_seg, overall_seg,
            ])

    # ── Write *_SessionCohort.csv ────────────────────────────────────────
    session_count = 0
    with open(session_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(SESSIONCOHORT_HEADER)

        for (uid, app), ops in sorted(session_ops.items()):
            n = len(ops)
            if n == 0:
                continue
            if n <= 5:
                cohort = "1-5 sessions"
            elif n <= 10:
                cohort = "6-10 sessions"
            elif n <= 20:
                cohort = "11-20 sessions"
            elif n <= 40:
                cohort = "21-40 sessions"
            elif n <= 60:
                cohort = "41-60 sessions"
            elif n <= 80:
                cohort = "61-80 sessions"
            else:
                cohort = "81+ sessions"
            writer.writerow([uid_original[uid], app, cohort])
            session_count += 1

    t_elapsed = time.perf_counter() - t_start

    if not quiet:
        print(f"[UserStats]     {total_users:,} users \u2192 {userstats_path.name} "
              f"({len(USERSTATS_HEADER)} columns)")
        print(f"[SessionCohort] {session_count:,} (user, app) pairs \u2192 {session_path.name}")
        print(f"[UserStats]     Elapsed: {t_elapsed:.2f}s")

    return total_users, session_count


# ═════════════════════════════════════════════════════════════════════════════
# CLI ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def main() -> None:
    parser = argparse.ArgumentParser(
        description=f"Purview M365 Usage Bundle Explosion Processor v{SCRIPT_VERSION} — "
        "Rollup-aggregated or event-level export of Purview audit log CSV for Power BI.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
examples (rollup — default, 80%%+ row reduction):
  python %(prog)s --input Purview_Export.csv
  python %(prog)s -i Purview_Export.csv --output-dir ./output

examples (event-level — v1-compatible 153-column output):
  python %(prog)s --mode event-level -i Purview_Export.csv
  python %(prog)s --mode event-level -i Purview_Export.csv --output-dir ./output

reconciliation (validate rollup correctness on a sample):
  python %(prog)s -i Purview_Export.csv --reconcile
""",
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to the input Purview audit log CSV file (must contain AuditData column).",
    )
    parser.add_argument(
        "--output-dir", "-o",
        default=None,
        help="Directory for output files. Default: same directory as the input file.",
    )
    parser.add_argument(
        "--mode", "-m",
        choices=["rollup", "event-level"],
        default="rollup",
        help="Processing mode. 'rollup' (default): 9-column aggregated output with 80%%+ row reduction. "
             "'event-level': v1-compatible 153-column output with one row per event.",
    )
    parser.add_argument(
        "--reconcile",
        action="store_true",
        default=False,
        help="Run sample-based reconciliation after processing to validate rollup correctness.",
    )
    parser.add_argument(
        "--prompt-filter",
        choices=["Prompt", "Response", "Both", "Null"],
        default=None,
        help="Filter Copilot messages: Prompt (user only), Response (AI only), Both (non-null), Null (null isPrompt).",
    )
    parser.add_argument(
        "--no-userstats",
        action="store_true",
        default=False,
        help="Skip generating *_UserStats.csv and *_SessionCohort.csv (rollup mode only).",
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        default=False,
        help="Suppress progress output (only errors are printed).",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {SCRIPT_VERSION}",
    )

    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    if not os.path.isfile(input_path):
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # ── Determine output directory & build filenames ─────────────────────
    stem = Path(input_path).stem
    output_dir = Path(os.path.abspath(args.output_dir)) if args.output_dir else Path(input_path).parent
    os.makedirs(output_dir, exist_ok=True)

    run_ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    if args.mode == "rollup":
        rollup_path = str(output_dir / f"{stem}_Rollup_{run_ts}.csv")
        userstats_path = str(output_dir / f"{stem}_UserStats_{run_ts}.csv")
        session_path = str(output_dir / f"{stem}_SessionCohort_{run_ts}.csv")
    else:
        rollup_path = str(output_dir / f"{stem}_Exploded_{run_ts}.csv")

    # ── Dispatch ─────────────────────────────────────────────────────────
    if args.mode == "rollup":
        stats = run_rollup(
            input_csv=input_path,
            output_csv=rollup_path,
            prompt_filter=args.prompt_filter,
            quiet=args.quiet,
        )
        exit_code = 1 if stats["parse_errors"] > stats["input_records"] * 0.1 else 0

        # ── UserStats & SessionCohort (derived from aggregated output) ───
        if not args.no_userstats:
            write_userstats_files(rollup_path, userstats_path, session_path, args.quiet)
    else:
        stats = run_explosion(
            input_csv=input_path,
            output_csv=rollup_path,
            prompt_filter=args.prompt_filter,
            quiet=args.quiet,
        )
        exit_code = 1 if stats["errors"] > 0 else 0

    # ── Optional reconciliation ──────────────────────────────────────────
    if args.reconcile:
        reconcile_passed = run_reconcile(
            input_csv=input_path,
            prompt_filter=args.prompt_filter,
            quiet=args.quiet,
        )
        if not reconcile_passed:
            exit_code = 1

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
