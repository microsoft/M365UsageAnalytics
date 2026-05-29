# M365 Usage Dashboard

<div align="center">

**User-level Microsoft 365 adoption and Copilot engagement reporting, powered by Purview audit logs.**

[![Built by Microsoft](https://img.shields.io/badge/Built%20by-Microsoft-0078d4?style=for-the-badge&logo=microsoft&logoColor=white)](https://github.com/microsoft/Analytics-Hub)
[![Analytics Hub](https://img.shields.io/badge/Analytics%20Hub-9%20Repositories-8661c5?style=for-the-badge&logo=github&logoColor=white)](https://github.com/microsoft/Analytics-Hub)

### 📥 [Click Here to Download All Files](https://github.com/microsoft/M365UsageAnalytics/archive/refs/heads/main.zip)

**Additional Resources:**
[Microsoft Purview Portal](https://purview.microsoft.com/) · [Microsoft Entra Admin Center](https://entra.microsoft.com/) · [PAX Purview Script](https://github.com/microsoft/PAX)

⭐ **Star this repository** to receive notifications about new template versions<br>
👀 **Watch** for updates and announcements

</div>

---

| | |
|---|---|
| **Data Source** | Microsoft Purview — Unified Audit Log |
| **Data Source** | Microsoft Entra ID — User Profiles |
| **Asset** | Power BI Template (`.pbit`) |
| **Output** | CSV files — Purview audit log export (processed) and Entra user/licensing export |

**[Quick Start ↓](#-quick-start)** &nbsp;·&nbsp; **[What This Dashboard Shows ↓](#-what-this-dashboard-shows)** &nbsp;·&nbsp; **[Report Pages ↓](#-report-pages-overview)** &nbsp;·&nbsp; **[Troubleshooting ↓](#-troubleshooting)**

---

> ⚠️ **Action Required: Microsoft Graph Audit API Permission Change (April 2026)**
>
> Microsoft introduced a new dedicated permission, `AuditLogsQuery.Read.All`, for the Microsoft Graph audit query API and began enforcing it across all tenants in April 2026.
>
> - The legacy `AuditLog.Read.All` permission is no longer sufficient to retrieve `CopilotInteraction` records.
> - Graph API calls with only the legacy permission will appear to succeed but return `0` records silently.
> - [PAX v1.10.9](https://github.com/microsoft/PAX) or later requests the correct scopes automatically. Upgrade earlier versions and grant admin consent for the new permission(s).

---

## 🚀 Quick Start

This section walks you through getting the dashboard running. There are **4 steps**, but the recommended path (PAX script) collapses them into **one command + opening Power BI**.

### Before you begin

- [ ] **Power BI Desktop** installed (June 2024 or later)
- [ ] **PowerShell 7+** installed (for PAX script)
- [ ] **Python 3.9+** installed (only needed for manual Purview exports)
- [ ] Required admin roles assigned — see [Roles & Permissions](#-roles--permissions)

> 💡 **First time? Start with 1 month of data.** A 1-month time window is the recommended starting point: it loads quickly, stays well under Purview's 50K/100K row export caps for most tenants, and produces meaningful tier rankings for an initial review. For long-running production reporting, widen the window to 90+ days once you've validated the pipeline end-to-end (more history → more stable percentile-based tier ranking).

> 📧 **Need your IT admin to export the data?** This pre-written email covers all required data sources, fields, roles, and steps — everything they need in one click.<br>
> **[📨 Email Prerequisites to Your IT Admin](mailto:?subject=Action%20Required%3A%20Data%20Export%20Fields%20Needed%20for%20M365%20Copilot%20Readiness%20Report%20%28Power%20BI%29&body=To%3A%20IT%20Admin%20%2F%20Global%20Admin%0ARe%3A%20M365%20Copilot%20Readiness%20Report%20%E2%80%93%20Power%20BI%20Report%20Setup%0A%0A%0AWHAT%20THIS%20REPORT%20DOES%0A%0AThe%20M365%20Copilot%20Readiness%20Report%20is%20a%20Power%20BI%20report%20that%20turns%20your%20M365%20Unified%20Audit%20Log%20into%20a%20prioritized%2C%20tiered%20view%20of%20which%20users%20are%20ready%20for%20Copilot%20licensing.%20It%20scores%20and%20ranks%20unlicensed%20users%20by%20M365%20app%20activity%20%28Teams%2C%20Outlook%2C%20Word%2C%20Excel%2C%20PowerPoint%29%2C%20classifies%20users%20into%20enablement%20strategy%20quadrants%2C%20and%20shows%20week-over-week%20engagement%20trends%20by%20workload.%20Designed%20for%20IT%20leaders%20making%20data-driven%20Copilot%20licensing%20and%20rollout%20decisions.%0A%0A%0ADATA%20SOURCES%20REQUIRED%0A%0A1.%20Microsoft%20Purview%20%E2%80%93%20M365%20Unified%20Audit%20Log%0A%20%20%20Export%3A%20Purview%20portal%20%28purview.microsoft.com%29%20-%3E%20Audit%20-%3E%20Search%20%28all%20activities%2C%2090%20day%20range%29%20-%3E%20Export%2C%20or%20PAX%20PowerShell%20script%20with%20-IncludeM365Usage%20switch%0A%20%20%20Format%3A%20CSV%0A%0A2.%20Microsoft%20Entra%20ID%20%E2%80%93%20User%20Details%0A%20%20%20Export%3A%20entra.microsoft.com%20-%3E%20Identity%20-%3E%20Users%20-%3E%20Download%20users%2C%20or%20PAX%20script%20with%20-IncludeUserInfo%20switch%0A%20%20%20Format%3A%20CSV%0A%0A%0AREQUIRED%20FIELDS%20%E2%80%94%20DO%20NOT%20REMOVE%0A%0AIMPORTANT%3A%20This%20report%20depends%20on%20specific%20operation%20types%20from%20the%20M365%20Unified%20Audit%20Log.%20If%20you%20are%20using%20the%20PAX%20script%2C%20run%20it%20with%20the%20-IncludeM365Usage%20flag%20%E2%80%94%20this%20ensures%20the%20correct%20workload%20activity%20types%20are%20captured.%20Do%20not%20filter%20out%20operation%20types%20or%20remove%20columns%20from%20the%20exported%20file%20before%20loading%20into%20Power%20BI.%20If%20running%20a%20manual%20Purview%20export%2C%20do%20not%20pre-process%20or%20re-save%20the%20file.%20Column%20order%20and%20formatting%20must%20be%20preserved%20exactly%20as%20exported.%0A%0APurview%20Unified%20Audit%20Log%20%E2%80%94%20Required%20Columns%3A%0ACreationDate%2C%20UserIds%2C%20Operations%2C%20AuditData.%0A%0APurview%20Audit%20Log%20%E2%80%94%20Required%20Operation%20Types%20%28within%20the%20export%29%3A%0AFileAccessed%2C%20FileModified%20%28SharePoint%2FOneDrive%29%2C%20MessageSent%20%28Teams%2FExchange%29%2C%20Teams%20meeting%20events%2C%20Word%20activity%20events%2C%20Excel%20activity%20events%2C%20PowerPoint%20activity%20events%2C%20OneNote%20activity%20events.%0A%0AMicrosoft%20Entra%20ID%20%E2%80%93%20User%20Export%3A%0AUserPrincipalName%2C%20displayName%2C%20Department%2C%20JobTitle%2C%20hasLicense%20%28or%20assignedLicenses%29.%0A%0ANote%3A%20The%20hasLicense%20column%20is%20critical%20for%20license-based%20insights.%20The%20PAX%20script%20adds%20this%20automatically.%20If%20exporting%20manually%20from%20Entra%2C%20you%20must%20populate%20this%20column%20with%20True%2FFalse%20based%20on%20whether%20each%20user%20has%20a%20Microsoft%20365%20Copilot%20SKU%20assigned.%20Without%20it%2C%20the%20Enablement%20Strategy%20quadrant%20and%20LP%20Score%20ranking%20will%20not%20distinguish%20licensed%20from%20unlicensed%20users.%0A%0A%0AINSIGHTS%20YOU%20WILL%20GAIN%0A%0A-%20Composite%20Licensing%20Priority%20%28LP%29%20score%20ranking%20every%20user%20by%20their%20M365%20app%20activity%20%E2%80%94%20ready%20for%20export%20and%20stakeholder%20review%0A-%20Copilot%20Enablement%20Strategy%3A%202x2%20quadrant%20classifying%20users%20as%20Enablement%20Targets%20%28high%20M365%20%2B%20low%20Copilot%29%2C%20Champions%2C%20AI-First%2C%20or%20Low%20Engagement%0A-%205-tier%20priority%20view%3A%20Critical%2C%20High%2C%20Medium%2C%20Promoter%2C%20Low%20%E2%80%94%20based%20on%20the%20gap%20between%20M365%20engagement%20and%20Copilot%20adoption%20per%20app%0A-%20Week-over-week%20trend%20lines%20across%20Teams%2C%20Outlook%2C%20Word%2C%20Excel%2C%20and%20PowerPoint%0A-%20User%20engagement%20segmentation%20%28normalized%20to%20active%20days%20per%20week%29%3A%20Daily%2C%20Frequent%2C%20Moderate%2C%20Light%2C%20Inactive%20%E2%80%94%20cross-referenced%20with%20Copilot%20license%20status%0A-%20Department-level%20engagement%20heatmaps%20and%20cold%20spots%0A%0A%0AROLES%20%26%20PERMISSIONS%20REQUIRED%0A%0AExport%20Purview%20Unified%20Audit%20Log%3A%20Audit%20Reader%20or%20Compliance%20Administrator%0AExport%20Entra%20user%20details%3A%20User%20Administrator%20or%20Global%20Reader%0ARun%20PAX%20script%20%28recommended%20for%20full%20coverage%29%3A%20Audit%20Reader%20%2B%20AuditLog.Read.All%2C%20User.Read.All%20%28Microsoft%20Graph%29%0A%0A%0ASOFTWARE%20REQUIREMENTS%0A%0A-%20Power%20BI%20Desktop%20%E2%80%94%20required%20to%20open%20the%20.pbip%20template%20file%20%28note%3A%20this%20report%20uses%20.pbip%20format%2C%20not%20.pbit%29%0A-%20PowerShell%205.1%2B%20%E2%80%94%20required%20only%20if%20using%20the%20PAX%20automated%20export%20script%0A-%20Access%20to%3A%20purview.microsoft.com%20%28or%20compliance.microsoft.com%29%2C%20entra.microsoft.com%0A%0A%0AImportant%20Notes%20on%20Export%20Size%3A%0AManual%20Purview%20exports%20are%20capped%20at%2050%2C000%20rows%20%28Audit%20Standard%29%20or%20100%2C000%20rows%20%28Audit%20Premium%29%20per%20search%20job.%20For%20tenants%20with%20high%20activity%20or%20date%20ranges%20over%2090%20days%2C%20use%20the%20PAX%20PowerShell%20script%20instead%20%E2%80%94%20it%20handles%20pagination%20automatically%20with%20no%20row%20limits%20and%20can%20be%20scheduled%20to%20run%20unattended%20on%20a%20recurring%20basis.%0A%0ARecommended%20Date%20Range%3A%0AA%20minimum%20of%2090%20days%20of%20audit%20data%20is%20required%20for%20meaningful%20tier%20ranking.%20Shorter%20ranges%20may%20classify%20all%20users%20as%20%22Developing%22%20due%20to%20insufficient%20activity%20history.)**

---

### Step 1. Export your data

You need two data exports: **(a)** the M365 Unified Audit Log from Purview and **(b)** user details from Entra ID. Pick one of the two paths below — A is recommended.

---

#### Path A — PAX script (recommended, one command does everything)

The [PAX script](https://github.com/microsoft/PAX) pulls audit data, processes it into the format Power BI needs, **and** exports Entra user/licensing data — all in one run. **If you can run PAX, skip Step 2 (processing) entirely and jump to Step 3.**

```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -IncludeM365Usage -Rollup -IncludeUserInfo -StartDate "2026-04-21" -EndDate "2026-05-21"
```

This produces **5 import-ready CSV files**:

| File | Contents |
|---|---|
| **Rollup** | Aggregated per-user / per-app / per-day event counts |
| **UserStats** | One row per user with pre-computed metrics, tiers, and engagement segments |
| **SessionCohort** | One row per (UserId, App) pair with session-count buckets |
| **SessionStats** | One row per (UserId, Date, AppHost) with Copilot prompt / session / response / agent counts (powers the **Copilot Prompts ✨** column) |
| **EntraUsers** | User profiles with department, job title, and Copilot license status |

> After the command finishes, note the five file paths shown in the console output — you'll need them in **Step 3**. Then skip directly to **[Step 3](#step-3-open-in-power-bi-desktop)**.

<details>
<summary><strong>PAX command variations</strong></summary>

<br>

**Interactive web login** (easiest — no app registration required):
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -IncludeM365Usage -Rollup -IncludeUserInfo -StartDate "2026-04-21" -EndDate "2026-05-21"
```

**Keep the raw Purview CSV alongside rollup files** (useful for ad-hoc analysis):
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -IncludeM365Usage -RollupPlusRaw -IncludeUserInfo -StartDate "2026-04-21" -EndDate "2026-05-21"
```

**App Registration** (for scheduled / unattended runs):
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -ClientId "<app-id>" -TenantId "<tenant-id>" -ClientSecret "<secret>" -IncludeM365Usage -Rollup -IncludeUserInfo -StartDate "2026-04-21" -EndDate "2026-05-21"
```

**Entra users only** (no audit data):
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -OnlyUserInfo
```

📖 See the [PAX repository](https://github.com/microsoft/PAX) for full documentation, authentication setup, and advanced options.

</details>

<details>
<summary><strong>Why PAX over manual exports?</strong></summary>

<br>

- **No row limits** — Manual Purview exports cap at 50K–100K rows. PAX runs parallel queries and combines results, routinely handling tens of millions of records.
- **No date-range splitting** — PAX automatically breaks your range into time slices, runs them simultaneously, and stitches results together.
- **Built-in post-processing** — With `-Rollup`, PAX produces the four import-ready CSVs directly. No separate processing step needed.
- **Entra + Purview in one run** — With `-IncludeUserInfo`, the Entra user export (including `hasLicense` column) is produced alongside the audit data.
- **Schedulable and unattended** — Can run on a schedule via Windows Task Scheduler or Azure Automation.
- **Automatically resumes if interrupted** — Picks up where it left off with no lost progress.

</details>

---

#### Path B — Manual 4-pull Purview export (recommended for medium / large tenants)

The 22 operations split across four narrower Purview searches by workload. Each pull stays well under the 50K / 100K row cap individually, and the processor in Step 2 combines them into one Rollup.

**Run four Purview searches**, each over the **same** date range. For each pull, follow these seven steps but paste only the activities listed in the table below for that pull. Save each export with a distinct filename.

1. Go to [purview.microsoft.com](https://purview.microsoft.com/) → **Audit** → **Search**
2. Set your **date range** (1 month recommended for the first run; expand later)
3. In **Activities**, paste the operation types for this pull from the table below — do **not** leave this blank
4. Leave **Users** blank to capture the full tenant
5. Click **Search** and wait for the job to complete
6. Click the completed job → **Export results** → **Download all results**
7. Do **not** re-save or pre-process the exported CSV — column order must be preserved exactly

| # | Pull | Activities (paste into Purview **Activities** filter) | Save as |
|---|---|---|---|
| 1 | **Files / SharePoint / OneDrive** | `FileAccessed, FileViewed, FilePreviewed, FileModified, FileDownloaded, FileUploaded` | `pull1_files.csv` |
| 2 | **Outlook / Mail** | `MailItemsAccessed, MailboxLogin, Send` | `pull2_outlook.csv` |
| 3 | **Teams (chat + meetings + sessions)** | `MessageSent, MessageRead, MessagesListed, ChatRetrieved, ChatCreated, MeetingParticipantJoined, MeetingStarted, MeetingEnded, MeetingParticipantDetail, MeetingDetail, TeamsSessionStarted` | `pull3_teams.csv` |
| 4 | **Copilot + Connected AI** *(paste into **Record types** filter — not Activities)* | `CopilotInteraction, ConnectedAIAppInteraction` | `pull4_copilot.csv` |

> ⚠️ **All four pulls must cover the same date range** — otherwise per-user-per-week math in the dashboard will be off.
>
> ℹ️ **Pull #4 uses Record types, not Activities.** `CopilotInteraction` and `ConnectedAIAppInteraction` are record types in Purview Audit Search; paste them into the **Record types** filter (the **Activities** filter will not find them).
>
> ℹ️ **Meeting export window:** extend the date range on pull #3 by **+1 day** beyond your analysis period — Purview batches meeting events up to 24 hours after the meeting ends.
>
> ℹ️ **Casing matters.** DAX `IN { ... }` filters are case-sensitive. Use the exact CamelCase shown above.

If any individual pull still returns the row cap, split that pull into weekly chunks and concatenate the resulting CSVs into a single file (per workload) before Step 2.

After exporting all four pulls, also complete **[Path D — Entra user export](#path-d--entra-user-export)**, then go to **[Step 2](#step-2-process-the-purview-exports)**.

---

<a id="path-d--entra-user-export"></a>
#### Path D — Entra user export

Required for **Path B** (PAX users already have this from Path A).

<details>
<summary><strong>Option 1 — Entra Admin Center (simplest)</strong></summary>

<br>

1. Go to [entra.microsoft.com](https://entra.microsoft.com/) → **Users** → **All users**
2. Click **Download users** → select **All users**
3. Ensure the export includes: `userPrincipalName`, `displayName`, `department`, `jobTitle`, `assignedLicenses`
4. Download as CSV

> ℹ️ The manual Entra export does not include a `hasLicense` column. The dashboard will load, but Copilot license-based insights won't be available until you add one. See **Option 2** below, or use the PAX script which adds it automatically.

</details>

<details>
<summary><strong>Option 2 — PowerShell with license detection (recommended, multi-SKU)</strong></summary>

<br>

<a id="entra-powershell-export"></a>

This snippet auto-discovers **every Copilot SKU currently provisioned in your tenant** (commercial, trial, GCC, Education, Frontline, Copilot Studio, etc.) and marks `hasLicense = TRUE` for any user holding **any** of them. Matching by **SKU GUID** is more resilient than substring-matching a part-number string, which Microsoft has renamed in the past.

```powershell
Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All"

# 1. Auto-discover every Copilot-related SKU in your tenant.
#    The -match 'Copilot' pattern catches: Microsoft_365_Copilot, Microsoft_365_Copilot_Trial,
#    Microsoft_365_Copilot_for_GCC, Microsoft_365_Copilot_for_Education_Faculty,
#    Microsoft_Copilot_Studio, CCIBOTS_PRIVPREV_VIRAL, Copilot_for_Sales, etc.
$copilotSkus = Get-MgSubscribedSku |
  Where-Object { $_.SkuPartNumber -match 'Copilot' -or $_.SkuPartNumber -match 'CCIBOTS' } |
  Select-Object SkuId, SkuPartNumber, ConsumedUnits

# 2. Print the matched SKUs so you can audit before exporting.
"Copilot-related SKUs found in tenant:" | Write-Host -ForegroundColor Cyan
$copilotSkus | Format-Table -AutoSize

# 3. (Optional) Scope down — uncomment to exclude industry Copilots from `hasLicense`.
#    Leave commented to count every Copilot product as licensed.
# $copilotSkus = $copilotSkus | Where-Object { $_.SkuPartNumber -notmatch 'Sales|Finance|Service' }

$copilotSkuIds = $copilotSkus.SkuId

# 4. Export every user with a hasLicense flag = TRUE if they hold any matched SKU.
Get-MgUser -All -Property userPrincipalName,displayName,department,jobTitle,assignedLicenses |
  Select-Object userPrincipalName, displayName, department, jobTitle,
    @{N='assignedLicenses'; E={ ($_.AssignedLicenses.SkuId) -join ',' }},
    @{N='hasLicense';      E={
        $userSkus = $_.AssignedLicenses.SkuId
        if ($null -eq $userSkus) { $false }
        else { ($userSkus | Where-Object { $copilotSkuIds -contains $_ }).Count -gt 0 }
    }} |
  Export-Csv -Path ".\EntraUsers.csv" -NoTypeInformation

"Export complete. Licensed user count:" | Write-Host -ForegroundColor Cyan
(Import-Csv .\EntraUsers.csv | Where-Object { $_.hasLicense -eq 'True' }).Count
```

> ℹ️ Re-export whenever there are significant directory changes (new hires, departures, license changes). Monthly is recommended.

</details>

Note the Entra CSV path — you'll use it as the `EntraUsers` parameter in **[Step 3](#step-3-open-in-power-bi-desktop)**.

---

#### Validating your Copilot license count

After loading the dashboard, sanity-check the **Has Copilot License = True** count on the **Copilot License Recommendations** page (or any page filtered by license) against the live tenant count.

**Quick cross-check** (no PowerShell required):

1. Go to the [Microsoft 365 admin center](https://admin.cloud.microsoft/) → expand **Reports**
2. Click **Usage**
3. Under **Reports**, click **Microsoft 365 Copilot**
4. Click the **Copilot** sub-item
5. Open the **Readiness** tab

The **Total prerequisite licenses** count shown there should match the **Has Copilot License = True** count from the dashboard (±a small delta for users added/removed between export time and now).

> ⚠️ Do **not** use this admin-center view as the dashboard's data source — it uses friendly product names instead of SKU part numbers and omits `department` / `jobTitle`. Use it only for visual verification.

**If counts don't match (the dashboard is too low):**

Your tenant may have a Copilot SKU that the matcher missed (rare — `-match 'Copilot'` catches almost everything). To diagnose and fix:

1. Run the discovery step from Option 2 above standalone to list every SKU currently assigned in your tenant:
   ```powershell
   Get-MgSubscribedSku | Where-Object { $_.ConsumedUnits -gt 0 } |
     Select-Object SkuPartNumber, SkuId, ConsumedUnits | Sort-Object SkuPartNumber
   ```
2. Identify any Copilot-related SKU not caught by `-match 'Copilot'` (for example, a custom-renamed SKU or a partner-specific bundle that includes Copilot).
3. Extend the filter in Option 2 to include it explicitly, e.g.:
   ```powershell
   $copilotSkus = Get-MgSubscribedSku | Where-Object {
       $_.SkuPartNumber -match 'Copilot' -or
       $_.SkuPartNumber -match 'CCIBOTS' -or
       $_.SkuPartNumber -eq 'YOUR_CUSTOM_SKU_PART_NUMBER'
   }
   ```
4. Re-run the export and refresh the PBIT.

> 📖 **SKU reference:** Microsoft's full list of license SKU part numbers and GUIDs is published at [Product names and service plan identifiers for licensing](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference). Use it to confirm SKU names if any look unfamiliar.

> ℹ️ **No PBIT change is needed** to add a new SKU — the matching logic lives in the PowerShell export, not in the report. Re-run the export with the updated SKU list and refresh; the dashboard picks up the change automatically.

---

### Step 2. Process the Purview exports

> **Skip this step if you used Path A (PAX with `-Rollup`).** The rollup CSVs are already import-ready. Go to [Step 3](#step-3-open-in-power-bi-desktop).

The raw Purview CSV(s) contain a nested `AuditData` JSON column that Power BI cannot import directly. The included processor (`Purview_M365_Usage_Bundle_Explosion_Processor_v2.6.0.py`) flattens it into four import-ready CSVs.

> 💡 **Don't have Python installed?** [Download the latest version at python.org](https://python.org).

**Run the processor on the four CSVs from Path B:**
```cmd
python scripts\Purview_M365_Usage_Bundle_Explosion_Processor_v2.6.0.py --files "pull1_files.csv" --outlook "pull2_outlook.csv" --teams "pull3_teams.csv" --copilot "pull4_copilot.csv"
```

> 💡 If you have a single raw Purview CSV (for example from a legacy PAX run without `-Rollup`), use `--pax "Purview_Export.csv"` instead.

Either invocation produces four output files (sharing a `_<YYYYMMDD_HHMMSS>` timestamp):

| File | Contents |
|---|---|
| **Rollup** (9 columns) | Aggregated event counts with min/max timestamps |
| **UserStats** (27 columns) | One row per user with pre-computed metrics, tier classifications, and engagement segments |
| **SessionCohort** (3 columns) | One row per (UserId, App) pair with a session-count bucket |
| **SessionStats** (7 columns) | One row per (UserId, Date, AppHost) with `PromptCount`, `SessionCount`, `ResponseCount`, `AgentSessionCount` — source for the **Copilot Prompts ✨** measure |

Note the output file paths — you'll need them in **Step 3**.

<details>
<summary><strong>Running from a Python interactive terminal instead</strong></summary>

<br>

```python
import subprocess
subprocess.run([
    "python", "scripts/Purview_M365_Usage_Bundle_Explosion_Processor_v2.6.0.py",
    "--pax", "Purview_Export.csv"
])
```

</details>

<details>
<summary><strong>Optional parameters and processing details</strong></summary>

<br>

| Parameter | Description |
|---|---|
| `--pax` | Single-export mode (raw PAX output) — one CSV in |
| `--teams` / `--outlook` / `--files` / `--copilot` | 4-pull mode (Path B) — one CSV per workload, combined into a single Rollup |
| `--output-dir` / `-o` | Directory for output files (default: input file's directory) |
| `--prompt-filter` | Filter Copilot messages: `Prompt`, `Response`, `Both`, or `Null` |
| `--skip-precompute` | Skip generating UserStats and SessionCohort files |
| `--reconcile` | Run sample-based reconciliation to validate processing correctness |
| `--debug-events` | Emit v1-compatible 153-column event-level CSV instead of rollup |
| `--quiet` / `-q` | Suppress progress output |

**What the script does:**
- Reads each row of the raw Purview CSV(s) and parses the `AuditData` JSON column
- Flattens nested objects and arrays (including Copilot event data with messages, contexts, and accessed resources)
- Aggregates the flattened events into rolled-up rows keyed by (UserId, CreationDate, Operation, Workload, SourceFileExtension, AppHost)
- Produces four output CSV files

> 💡 **Performance tip:** Install [orjson](https://pypi.org/project/orjson/) (`pip install orjson`) for 5–10× faster JSON parsing. The script falls back to Python's built-in parser if orjson is not available.

</details>

---

<a id="step-3-open-in-power-bi-desktop"></a>
### Step 3. Open in Power BI Desktop

1. Open **Power BI Desktop** → **File** → **Open report** → **Browse** → select the `.pbit` template from this folder
2. Go to **Home** → **Transform data** → **Edit parameters**
3. Set the five file paths:

   | Parameter | Value |
   |---|---|
   | `M365 Rollup Data` | Full path to the **Rollup** CSV |
   | `M365 User Stats Data` | Full path to the **UserStats** CSV |
   | `M365 Session Cohort Data` | Full path to the **SessionCohort** CSV |
   | `M365 Session Stats Data` | Full path to the **SessionStats** CSV (powers the **Copilot Prompts ✨** column) |
   | `Entra Users Data` | Full path to your Entra user details CSV |

4. Click **OK** → **Apply changes**
5. Click **Refresh** on the Home ribbon — allow several minutes on first load with large datasets
6. **Save as `.pbix`** — this avoids re-entering parameters on every open

---

## 🔄 Refreshing Data Over Time

1. Get fresh CSV exports (re-run PAX with updated dates, or re-export manually)
2. Process the Purview CSV if needed (skip if using PAX with `-Rollup`)
3. Overwrite the CSV files at the same paths your report points to
4. Open the saved `.pbix` and click **Refresh**

> The dashboard does not connect live to Purview or Entra — all data is point-in-time CSV exports.

---

## 📊 What This Dashboard Shows

<details>
<summary><strong>Expand for full description</strong></summary>

<br>

Copilot licensing decisions shouldn't be made on gut feel or org chart. This report turns your existing M365 Unified Audit Log into a tiered, prioritized view of user readiness — so you deploy Copilot licenses to the users most likely to get value from day one.

| Report Page | What You Can Answer |
|---|---|
| **M365 App Usage Report** | What is our overall M365 adoption rate? Which apps are driving the most engagement? |
| **M365 Usage Trends** | How is activity trending week over week? Which apps dominate the workload mix? |
| **Copilot License Recommendations** | Who should get a Copilot license first based on weighted M365 usage? |
| **Copilot Enablement Strategy** | Where are the biggest gaps between M365 usage and Copilot adoption? Who are Champions vs. Enablement Targets? |
| **Glossary and Metric Definitions** | Definitions for all metrics, tiers, engagement segments, and scoring methodology. |
| **M365 Usage Activity** | How are users distributed across engagement segments? Are Copilot-licensed users more active? |
| **Enablement Strategy — Priority Table** | Which users need training most urgently? What is each user's recommended next action? |
| **M365 Copilot Licensing Strategy** | Which users should be licensed in each wave? How do users rank across all M365 apps? |

</details>

---

## 🖥️ Report Pages Overview

<details>
<summary><strong>Expand to view all 8 report pages</strong></summary>

<br>

The dashboard includes **8 interactive report pages**. See the [Interpretation Guide](M365%20Usage%20Dashboard%20-%20Interpretation%20Guide.pdf) for a detailed walkthrough of each page.

[![Report Pages Slideshow](images/report-pages-carousel.gif)](images/report-pages-carousel.gif)
*💡 Expand any report page section below to view a full-size still screenshot.*

---

<details>
<summary><strong>1. M365 App Usage Report</strong></summary>

Your landing page — six headline KPI cards (total M365 users, Teams activity, email activity, document activity, Copilot users, and departments tracked), a stacked bar chart breaking down total events by app and action type, and a Report Highlights panel linking to the dashboard's four main analytical sections.

[![M365 App Usage Report](images/1%20-%20M365%20App%20Usage%20Report.png)](images/1%20-%20M365%20App%20Usage%20Report.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>2. M365 Usage Trends</strong></summary>

Headline KPIs for total users, active users, all-app actions, and average actions per user. A ranked bar chart shows which apps drive the most total activity, a donut chart highlights each app's proportional share, and a week-over-week trend line tracks average actions per app over time. A tier distribution matrix segments users into Top 10%, 10–25%, 25–50%, and Bottom 50% per application.

[![M365 Usage Trends](images/2%20-%20M365%20Usage%20Trends.png)](images/2%20-%20M365%20Usage%20Trends.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>3. Copilot License Recommendations</strong></summary>

Every user ranked by a weighted composite score (0–100) blending their percentile across selected M365 apps. Choose from four profile presets — Balanced, Collaboration Focus, Content Creation, or Custom — to model different licensing scenarios. Users are classified into action categories: License First (≥90th percentile), License Next (75th–89th), Potential (50th–74th), and Developing (<50th). A Customize Weight panel lets you manually adjust per-app weights.

[![Copilot License Recommendations](images/3%20-%20Copilot%20License%20Recommendations.png)](images/3%20-%20Copilot%20License%20Recommendations.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>4. Copilot Enablement Strategy</strong></summary>

A 2×2 quadrant model classifying every user into four segments: **Enablement Targets** (high M365 / low Copilot — best training candidates), **Champions** (high on both — peer advocates), **AI-First** (low M365 / high Copilot), and **Low Engagement** (low on both). Filter by app to see workflow-specific quadrants, and sort by the Enablement Gap column to find users with the largest gap between M365 activity and Copilot adoption.

[![Copilot Enablement Strategy](images/4%20-%20Copilot%20Enablement%20Strategy.png)](images/4%20-%20Copilot%20Enablement%20Strategy.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>5. Glossary and Metric Definitions</strong></summary>

In-report reference covering every metric, quadrant, tier label, engagement segment, and scoring methodology used across the dashboard. Includes definitions for active days, app tiers, composite scores, enablement gaps, and action categories so the report can be shared broadly without external documentation.

[![Glossary and Metric Definitions](images/5%20-%20Glossary%20and%20Metric%20Definitions.png)](images/5%20-%20Glossary%20and%20Metric%20Definitions.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>6. M365 Usage Activity</strong></summary>

Four KPI cards covering total users, M365 actions, average actions per week, and average active days. A bar chart ranks apps by average active days per user per week, and a segmentation chart groups users by engagement level — **Daily** (5+ days/week), **Frequent** (3–4), **Moderate** (1–2), **Light** (<1), and **Inactive** (0). Thresholds are computed by the processor as `active_days × 7 ÷ window_days`, where `window_days` is the calendar span of your Purview pull, so the labels mean the same thing whether the data covers 8 days or 6 months. A comparison chart shows whether Copilot-licensed users are more active than unlicensed users.

[![M365 Usage Activity](images/6%20-%20M365%20Usage%20Activity.png)](images/6%20-%20M365%20Usage%20Activity.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>7. Copilot Enablement Strategy — Priority Table</strong></summary>

Each user assigned a Priority level (Critical, High, Medium, Promoter, Low) and a recommended Action (Immediate Training, Train Next, Advanced Training, Monitor) based on the ratio of M365 activity to Copilot usage. App Tier and Copilot Tier columns show each user's relative standing. Filter by app and priority level to generate targeted training outreach lists.

[![Copilot Enablement Strategy — Priority Table](images/7%20-%20Copilot%20Enablement%20Strategy%20-%20Priority%20Table.png)](images/7%20-%20Copilot%20Enablement%20Strategy%20-%20Priority%20Table.png)
*Click image to enlarge*

</details>

<details>
<summary><strong>8. M365 Copilot Licensing Strategy</strong></summary>

A tier-based licensing planner with four waves: **Prioritize** (top M365 users — license immediately), **License Next** (strong candidates for the next quarter), **Enablement** (moderate usage — train before licensing), and **Monitor** (low activity — revisit later). Color-coded cells show each user's tier in Teams, Outlook, Word, Excel, and PowerPoint. Filter by action tier to export ready-made licensing request lists.

[![M365 Copilot Licensing Strategy](images/8%20-%20M365%20Copilot%20Licensing%20Strategy.png)](images/8%20-%20M365%20Copilot%20Licensing%20Strategy.png)
*Click image to enlarge*

</details>

</details>

---

## ⚖️ How Weighting Works

<details>
<summary><strong>Expand for scoring methodology</strong></summary>

<br>

The **Copilot License Recommendations** page uses a composite score to rank every user from 0–100.

#### Score Calculation

Each user receives a percentile rank within each M365 app relative to all other users. A score of 80 means the user is more active than 80% of peers in that app. The composite score is the weighted average of these per-app percentiles.

#### Profile Presets

| Profile | Weighting Approach | Best For |
|---|---|---|
| **Balanced** | Equal weights across Teams, Outlook, Word, Excel, PowerPoint | General licensing decisions |
| **Collaboration Focus** | Heavy weight on Teams and Outlook | Copilot rollouts prioritizing meeting and email productivity |
| **Content Creation** | Heavy weight on Word, Excel, PowerPoint | Document workflow productivity goals |
| **Custom** | You set the weight for each app manually (weights do not need to sum to 100) | Organization-specific scenarios |

#### Action Category Thresholds

| Category | Percentile | Recommended Action |
|---|---|---|
| **License First** | ≥ 90th | License immediately — highest M365 engagement |
| **License Next** | 75th–89th | Strong candidate for next licensing wave |
| **Potential** | 50th–74th | Provide enablement before licensing |
| **Developing** | < 50th | Focus on foundational M365 engagement first |

> 💡 **Tip:** Changing the profile instantly re-ranks all users. Try running Balanced, Collaboration Focus, and Content Creation profiles and compare which users appear in License First across all three — those are your safest first-wave candidates.

[![How Weighting Works](images/How-Weighting-Works.png)](images/How-Weighting-Works.png)
*Click image to enlarge*

</details>

---

## 🔑 Roles & Permissions

<details>
<summary><strong>Expand for full permission tables</strong></summary>

<br>

### Purview Audit Log — Manual Export

| Role | Where Assigned | Notes |
|---|---|---|
| `View-Only Audit Logs` | Microsoft Purview compliance portal → Roles | Minimum required — read-only access to audit search and export |
| `Audit Logs` | Microsoft Purview compliance portal → Roles | Read and configure access to audit search |
| `Compliance Administrator` | Microsoft Entra ID → Roles | Includes audit log access |
| `Global Administrator` | Microsoft Entra ID → Roles | Full access — use only if other roles aren't available |

### Purview Audit Log — PAX / Graph API

| Permission | Purpose | When required (Graph API) | Delegated | AppRegistration | EOM |
|---|---|---|---|---|---|
| `AuditLogsQuery.Read.All` | Umbrella permission for Graph audit query API — covers `CopilotInteraction` | Always except `-OnlyUserInfo` | ✅ | ✅ | N/A |
| `AuditLogsQuery-Exchange.Read.All` | Exchange Online audit logs | `-IncludeM365Usage` | ✅ | ✅ | N/A |
| `AuditLogsQuery-OneDrive.Read.All` | OneDrive audit logs | `-IncludeM365Usage` | ✅ | ✅ | N/A |
| `AuditLogsQuery-SharePoint.Read.All` | SharePoint Online audit logs | `-IncludeM365Usage` | ✅ | ✅ | N/A |
| `User.Read.All` | Entra user directory, MAC licensing | `-IncludeUserInfo`, `-OnlyUserInfo`, or `-GroupNames` | ✅ | ✅ | N/A |
| `Organization.Read.All` | Tenant / organization context, license metadata | `-IncludeUserInfo` or `-OnlyUserInfo` | ✅ | ✅ | N/A |
| `GroupMember.Read.All` | Group lookup and membership expansion | `-GroupNames` | ✅ | ✅ | N/A |
| `Purview Audit Reader` | Purview UI / EOM | EOM only | ❌ | ❌ | ✅ |

> **📚 Reference:** [PAX Script Documentation](https://github.com/microsoft/PAX) · [Create auditLogQuery — Permissions](https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries#permissions) · [Get auditLogQuery — Permissions](https://learn.microsoft.com/en-us/graph/api/security-auditlogquery-get#permissions)

### Entra ID User & Licensing Data

Manual export from the Entra admin center does not rely on Graph API permissions. If you use PAX or another Graph-based tool, grant:

| Permission | Purpose | When required |
|---|---|---|
| `User.Read.All` | Read Entra user profiles and MAC licensing data | `-IncludeUserInfo`, `-OnlyUserInfo`, or `-GroupNames` |
| `Organization.Read.All` | Read tenant context and license SKU metadata | `-IncludeUserInfo` or `-OnlyUserInfo` |
| `GroupMember.Read.All` | Expand group names and group membership | `-GroupNames` |

### Audit Log Retention

| License | Default Retention | Maximum |
|---|---|---|
| Audit (Standard) — E3 / A3 | 180 days | 180 days |
| Audit (Premium) — E5 / A5 | 1 year (Exchange, SharePoint, OneDrive, Entra); 180 days (all other) | 1 year |
| Audit (Premium) — E5 / A5 + 10-Year Add-On | Same defaults as E5 | Up to 10 years (via custom retention policy) |

> **Note:** The default retention for Audit (Standard) changed from 90 days to 180 days on October 17, 2023. Confirm your license tier before pulling data.

</details>

---

## 📋 Required Operation Types

<details>
<summary><strong>Expand for categorized operation types table</strong></summary>

<br>

This is the **minimum** set of operation types the dashboard's DAX measures and Power Query code actually read. It is validated against every visual, measure, and M-query expression in the PBIT (including the Copilot License Optimizer / agents page).

| Category | Operations |
|---|---|
| **Copilot — Office in-product** | CopilotInteraction |
| **Copilot — Agents / Connected AI apps** | ConnectedAIAppInteraction |
| **Outlook / Exchange** | MailItemsAccessed, MailboxLogin, Send |
| **SharePoint / OneDrive — Files** | FileAccessed, FileViewed, FilePreviewed, FileModified, FileDownloaded, FileUploaded |
| **Teams — Messaging** | MessageSent, MessageRead, MessagesListed, ChatRetrieved, ChatCreated |
| **Teams — Meetings** | MeetingParticipantJoined, MeetingStarted, MeetingEnded, MeetingParticipantDetail, MeetingDetail |
| **Teams — Sessions** | TeamsSessionStarted |

> ℹ️ **`ConnectedAIAppInteraction`** powers the Copilot License Optimizer agents telemetry (`IsAgentInteraction`, `AgentId`, `AgentName`, `ContextType`). It is required even though it appears in some measures only as an exclusion filter — without the rows in the export, the inclusion measures return zero for Connected AI apps.
>
> ℹ️ **Meeting events** are batched by Purview up to 24 hours after the meeting ends. When pulling meeting-related operations, extend your export window by **+1 day** beyond the period you want to analyze.
>
> ℹ️ **Casing is significant.** The DAX `IN { ... }` operator is case-sensitive. Use the exact CamelCase shown above. Lowercase variants (e.g. `mailitemsaccessed`) will silently drop rows.

**Record types** (for API-level filtering): ExchangeItem, SharePointFileOperation, OneDrive, MicrosoftTeams, CopilotInteraction, AIAppInteraction

</details>

---

## ⚠️ Known Issues & Workarounds

The following are known limitations in the current PBIT release. None block successful loading or core insights, but power users should be aware of them.

<details>
<summary><strong>1. Entra portal export uses friendly column headers (e.g. "User principal name")</strong></summary>

<br>

**Symptom:** Manual Entra portal exports come with display-friendly column headers (`User principal name`, `Display name`, `Job title`, `Assigned licenses`) rather than the Graph API camelCase names (`userPrincipalName`, `displayName`, `jobTitle`, `assignedLicenses`). The Power Query `HeaderMap` step in the EntraUsers table normalizes the five most common variants automatically, but uncommon header variants will leave a few columns unmapped.

**Workaround:** Before loading, rename the headers in Excel / your editor to match the Graph API names:

| Portal header | Required name |
|---|---|
| `User principal name` | `userPrincipalName` |
| `Display name` | `displayName` |
| `Job title` | `jobTitle` |
| `Department` | `department` |
| `Assigned licenses` | `assignedLicenses` |

Alternatively, use the [PAX script](https://github.com/microsoft/PAX) with `-IncludeUserInfo` — it emits the canonical camelCase header set directly.

</details>

<details>
<summary><strong>2. Manual Entra exports require a populated <code>hasLicense</code> column</strong></summary>

<br>

**Symptom:** All license-based visuals (License Recommendations page, LP percentile measures, Enablement Strategy quadrant) appear blank or undifferentiated when an Entra export is loaded that does **not** include a populated `hasLicense` column.

**Cause:** The `Has Copilot License` calculated column reads `EntraUsers[hasLicense] = TRUE` directly. The PAX script populates this column automatically; manual `entra.microsoft.com` exports do not.

**Fix:** Use either the PAX script with `-IncludeUserInfo`, or the multi-SKU PowerShell snippet in **[Quick Start → Step 1 → Path D, Option 2](#step-1--export-the-required-data)**, which auto-discovers every Copilot SKU in the tenant and marks `hasLicense = TRUE` for any user holding any of them.

</details>

<details>
<summary><strong>3. <code>createdDateTime</code> column typed as string in the model</strong></summary>

<br>

**Symptom:** In the `EntraUsers` table, the `createdDateTime` column is loaded as text rather than a datetime.

**Impact:** Zero impact on any current visual or measure — `createdDateTime` is not referenced anywhere in the dashboard.

**Workaround:** No action needed today. If you build a custom visual that filters or sorts by account creation date, change the column type to `Date/Time` in Power Query first.

</details>

---

## 🛠️ Troubleshooting

<details>
<summary><strong>Errors on load — type conversion failures</strong></summary>

**Symptom:** Thousands of errors shown after refresh, visible in the Power Query diagnostics pane.

**Cause:** The Purview CSV writes missing values as the literal text string `"null"` rather than leaving cells blank.

**Fix:** Already handled in the M Code — the `Replace Null Strings` step converts `"null"` and `"None"` to actual nulls before type casting. If you still see errors, check that your CSV was processed through the processor script and that you're pointing at the correct Rollup CSV.

</details>

<details>
<summary><strong>Date errors — "month 13" or date parse failures</strong></summary>

**Symptom:** Date columns fail to load; error messages reference invalid month values.

**Cause:** The Purview CSV uses M/D/YYYY date format (US locale). Non-US locale Power BI installations parse the day as a month.

**Fix:** Already handled — the `Changed Type` M Code step includes `"en-US"` as an explicit locale argument. If dates still fail, confirm your Power BI Desktop regional settings match the CSV format.

</details>

<details>
<summary><strong>Users not joining — blank department/jobTitle on all rows</strong></summary>

**Symptom:** Department and job title slicers are empty; all user-attribute measures return blank.

**Cause:** The join between `SecDemoM365Usage[UserId]` and `SecDemEntraUsers[userPrincipalName]` is not matching — usually due to case differences or alias vs. canonical UPN.

**Fix:** Open Power Query, preview both columns, and compare a sample of values. Ensure both exports come from the same tenant and the Entra export uses `userPrincipalName` (not `mail` or `id`).

</details>

<details>
<summary><strong>Template prompts for file paths every time I open it</strong></summary>

**Symptom:** Every open of the `.pbit` file re-prompts for CSV paths.

**Cause:** Expected behavior for `.pbit` templates.

**Fix:** After the first load, save as `.pbix`. Subsequent opens will use the saved parameter values.

</details>

<details>
<summary><strong>Additional troubleshooting reference</strong></summary>

| Symptom | Cause | Fix |
|---|---|---|
| Blank matrix, no rows | Hardcoded slicer default filtering to a non-existent value | Remove the `filter` property from `objects.general[0].properties` in the slicer's `visual.json` |
| Cards showing 0 | EntraData filtered to zero rows by a slicer default | Same fix — remove hardcoded department or group filter |
| LP table blank | Row field set to `M365Usage.UserId` instead of `EntraData.userPrincipalName` | Change the row field in the visual's `queryState.Rows` |
| LP table blank after row field fix | Wrong filter measure for the page type | Replace `App Selection.Show User Based on App Tier Filter` with `LP Measures.LP Action Filter Match` |
| All users showing "Developing" | No M365 activity in the export date range | Widen the export to at least 90 days |
| Tier ranking incorrect | Zero-activity users included in percentile calculation | Confirm `PERCENTILEX` filters to `[Measure] > 0` before ranking |
| Refresh type error on JSON column | Mixed integer/string fields in some audit events | Use `try Int64.From(...) otherwise null` in the M Code extraction step |
| Report fails to open (relationship error) | Removed columns that had auto-date relationships | Remove orphaned relationship blocks and LocalDateTable refs from `relationships.tmdl` and `model.tmdl` |
| Slicer shows blank/garbage first item | Header row not removed after `Binary.Decompress` | Add `Table.Skip(#"Renamed Columns", 1)` after renaming columns |

</details>

---

## 📋 Next Steps

<details>
<summary><strong>Publish / Distribute</strong></summary>

- Save your file as a `.pbix` after setup.
- Publish to a Power BI workspace via **Home** → **Publish**.
- Open the **Semantic Model** settings and configure **Data source credentials** to point to the CSV file locations.
- For scheduled refresh, store the CSV files in SharePoint or OneDrive.

</details>

<details>
<summary><strong>Interpretation & Action Planning</strong></summary>

Use the report pages in this order to tell a complete Copilot readiness story:

1. **Executive Summary** — Overall tenant snapshot
2. **M365 Usage Trends** — Week-over-week engagement by app
3. **Copilot License Recommendations** — Ranked candidates with adjustable weights
4. **Copilot Enablement Strategy** — 2×2 quadrant: Champions, Enablement Targets, AI-First, Low Engagement
5. **Glossary & Definitions** — Tier definitions, scoring methodology
6. **M365 Usage Activity** — Baseline engagement segments, active day comparisons
7. **Enablement Strategy — Priority Table** — Critical/High/Medium/Promoter/Low with recommended actions
8. **Copilot Licensing Strategy** — Tier-based wave planner: Prioritize, License Next, Enablement, Monitor

</details>

<details>
<summary><strong>Monitor with Automatic Refresh</strong></summary>

- Publish to Power BI Service and configure scheduled refresh (weekly recommended).
- Re-run your Purview export regularly — if using PAX with `-IncludeM365Usage -Rollup`, the entire pipeline runs in one command, ideal for scheduling via Windows Task Scheduler or Azure Automation.
- Overwrite the previous CSV files at the same paths; the report picks up new data on next refresh.
- Track tier distribution changes over time — users moving from Bottom 50% into Top 25% indicates Copilot adoption success.

</details>

---

## 💬 Feedback

Managed and released by the Microsoft Copilot Growth ROI Advisory Team. Reach out to [copilot-roi-advisory-team-gh@microsoft.com](mailto:copilot-roi-advisory-team-gh@microsoft.com) with any feedback.

For questions about the Purview export scripts, see the [PAX repository](https://github.com/microsoft/PAX).

---

> **Data handling reminder:** This dashboard processes user-level activity data from your Microsoft 365 tenant. Treat all exported CSVs as sensitive. Do not store them in shared locations, commit them to source control, or distribute them outside your organization's data governance boundaries.
