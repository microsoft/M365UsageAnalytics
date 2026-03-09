#!/usr/bin/env python3
"""
Purview M365 Usage Bundle Explosion Processor v1.0.0
=====================================================
Standalone Python script that explodes/flattens Purview audit log CSV exports
containing AuditData JSON columns into fully denormalized rows for Power BI ingestion.

Replicates the explosion logic from PAX_Purview_Audit_Log_Processor (PowerShell) in
native Python for maximum performance.

Requirements:
    Python 3.9+
    pip install orjson   (OPTIONAL - 5-10x faster JSON parsing; falls back to stdlib json)

Usage:
    python Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.py --input <CSV> [--output <CSV>] [--prompt-filter Prompt|Response|Both|Null] [--workers N] [--chunk-size N] [--quiet]

Examples:
    python Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.py -i Purview_Export.csv
    python Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.py -i Purview_Export.csv -o Exploded.csv --workers 4
    python Purview_M365_Usage_Bundle_Explosion_Processor_v1.0.0.py -i Purview_Export.csv --prompt-filter Prompt --quiet

Author:  PAX Copilot Analytics Team
Version: 1.0.0
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys
import time
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

SCRIPT_VERSION = "1.0.0"

EXPLOSION_PER_RECORD_ROW_CAP = 1000
STREAMING_CHUNK_SIZE = 5000

# 121 M365 Usage Activity Types (matches PAX $m365UsageActivityBundle, excluding UserLoggedIn)
M365_USAGE_ACTIVITY_BUNDLE: set[str] = {
    # Exchange/Email
    "MailboxLogin", "MailItemsAccessed", "Send", "SendOnBehalf", "SoftDelete",
    "HardDelete", "MoveToDeletedItems", "CopyToFolder",
    # SharePoint/OneDrive - Files
    "FileAccessed", "FileDownloaded", "FileUploaded", "FileModified", "FileDeleted",
    "FileMoved", "FileCheckedIn", "FileCheckedOut", "FileRecycled", "FileRestored",
    "FileVersionsAllDeleted",
    # SharePoint/OneDrive - Sharing
    "SharingSet", "SharingInvitationCreated", "SharingInvitationAccepted",
    "SharedLinkCreated", "SharingRevoked", "AddedToSecureLink", "RemovedFromSecureLink",
    "SecureLinkUsed",
    # Groups/Unified Groups
    "AddMemberToUnifiedGroup", "RemoveMemberFromUnifiedGroup",
    # Teams - Team/Channel management
    "TeamCreated", "TeamDeleted", "TeamArchived", "TeamSettingChanged",
    "TeamMemberAdded", "TeamMemberRemoved", "MemberAdded", "MemberRemoved",
    "MemberRoleChanged", "ChannelAdded", "ChannelDeleted", "ChannelSettingChanged",
    "ChannelOwnerResponded", "ChannelMessageSent", "ChannelMessageDeleted",
    "BotAddedToTeam", "BotRemovedFromTeam", "TabAdded", "TabRemoved", "TabUpdated",
    "ConnectorAdded", "ConnectorRemoved", "ConnectorUpdated",
    # Teams - Chat/Messaging
    "TeamsSessionStarted", "ChatCreated", "ChatRetrieved", "ChatUpdated",
    "MessageSent", "MessageRead", "MessageDeleted", "MessageUpdated", "MessagesListed",
    "MessageCreation", "MessageCreatedHasLink", "MessageEditedHasLink",
    "MessageHostedContentRead", "MessageHostedContentsListed", "SensitiveContentShared",
    # Teams - Meeting lifecycle
    "MeetingCreated", "MeetingUpdated", "MeetingDeleted", "MeetingStarted",
    "MeetingEnded", "MeetingParticipantJoined", "MeetingParticipantLeft",
    "MeetingParticipantRoleChanged", "MeetingRecordingStarted", "MeetingRecordingEnded",
    "MeetingDetail", "MeetingParticipantDetail", "LiveNotesUpdate", "AINotesUpdate",
    "RecordingExported", "TranscriptsExported",
    # Teams - Apps/Approvals
    "AppInstalled", "AppUpgraded", "AppUninstalled", "CreatedApproval",
    "ApprovedRequest", "RejectedApprovalRequest", "CanceledApprovalRequest",
    # Office apps
    "Create", "Edit", "Open", "Save", "Print",
    # Microsoft Forms
    "CreateForm", "EditForm", "DeleteForm", "ViewForm", "CreateResponse",
    "SubmitResponse", "ViewResponse", "DeleteResponse",
    # Microsoft Stream
    "StreamModified", "StreamViewed", "StreamDeleted", "StreamDownloaded",
    # Planner
    "PlanCreated", "PlanDeleted", "PlanModified", "TaskCreated", "TaskDeleted",
    "TaskModified", "TaskAssigned", "TaskCompleted",
    # Power Apps
    "LaunchedApp", "CreatedApp", "EditedApp", "DeletedApp", "PublishedApp",
    # Copilot
    "CopilotInteraction",
}

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

# Set for fast membership checks
_UNIFIED_HEADER_SET: frozenset[str] = frozenset(M365_UNIFIED_HEADER)

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


def is_scalar(val: Any) -> bool:
    """Check if value is a scalar (str, int, float, bool, or None)."""
    return val is None or isinstance(val, (str, int, float, bool))


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
# CORE FLATTENING FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════




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
# CLI ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def main() -> None:
    parser = argparse.ArgumentParser(
        description=f"Purview M365 Usage Bundle Explosion Processor v{SCRIPT_VERSION} — "
        "Explodes/flattens Purview audit log CSV exports for Power BI.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
examples:
  python %(prog)s --input Purview_Export.csv
  python %(prog)s -i Purview_Export.csv -o Exploded.csv --workers 4
  python %(prog)s -i Purview_Export.csv --prompt-filter Prompt --workers 4
""",
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to the input Purview audit log CSV file (must contain AuditData column).",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Path for the output exploded CSV. Default: <input_stem>_Exploded.csv",
    )
    parser.add_argument(
        "--prompt-filter",
        choices=["Prompt", "Response", "Both", "Null"],
        default=None,
        help="Filter Copilot messages: Prompt (user only), Response (AI only), Both (non-null), Null (null isPrompt).",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=0,
        help="Number of parallel workers. Default: min(cpu_count, 8). Use 1 to disable parallelism.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=STREAMING_CHUNK_SIZE,
        help=f"Number of CSV rows per processing chunk. Default: {STREAMING_CHUNK_SIZE}.",
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

    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        stem = Path(input_path).stem
        parent = Path(input_path).parent
        output_path = str(parent / f"{stem}_Exploded.csv")

    stats = run_explosion(
        input_csv=input_path,
        output_csv=output_path,
        prompt_filter=args.prompt_filter,
        workers=args.workers,
        chunk_size=args.chunk_size,
        quiet=args.quiet,
    )

    sys.exit(1 if stats["errors"] > 0 else 0)


if __name__ == "__main__":
    main()
