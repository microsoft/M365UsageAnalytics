# M365 Copilot Readiness Report

<div align="center">

![Current Version](https://img.shields.io/badge/version-1-blue)

**Identify which users are ready for Microsoft Copilot licensing—and build a prioritized, data-driven adoption strategy from your own M365 Unified Audit Log.**

### 📥 [Click Here to Download All Files](https://github.com/microsoft/M365UsageAnalytics/archive/refs/heads/main.zip)

**Related Templates & Tools:**

[![Super User Impact](https://img.shields.io/badge/Report-Super%20User%20Impact-003087)](https://aka.ms/superuserimpact)
[![AI-in-One Dashboard](https://img.shields.io/badge/Report-AI--in--One%20Dashboard-teal)](https://github.com/microsoft/AI-in-One-Dashboard)
[![GitHub Copilot Impact](https://img.shields.io/badge/Report-GitHub%20Copilot%20Impact-purple)](https://github.com/microsoft/GitHubCopilotImpact)
[![Chat Intelligence](https://img.shields.io/badge/Report-Chat%20Intelligence-orange)](https://github.com/microsoft/CopilotChatAnalytics)
[![PBI to Exec Deck](https://img.shields.io/badge/Tool-PBI%20to%20Exec%20Deck-red)](https://github.com/shailendrahegde/pbi-to-exec-deck)

**Additional Resources:**
[Microsoft Purview Portal](https://purview.microsoft.com/), [Microsoft Entra Admin Center](https://entra.microsoft.com/), [PAX Purview Script](https://github.com/microsoft/PAX)

⭐ **Star this repository** to receive notifications about new template versions
👀 **Watch** for updates and announcements

</div>

---

<details open>
<summary><strong>📊 Why Use This Report & Insights You Can Explore</strong></summary>

<br>

Copilot licensing decisions shouldn't be made on gut feel or org chart. This report turns your existing M365 Unified Audit Log into a tiered, prioritized view of user readiness — so you deploy Copilot licenses to the users most likely to get value from day one.

**Executive Overview:**
How active is the tenant across M365 workloads? How many users are Copilot-licensed vs. unlicensed?

**Copilot Enablement Strategy:**
Which users are highly active in M365 but underusing Copilot? Who are your Champions? Where are the biggest gaps between M365 engagement and Copilot adoption?

**Licensing Priority:**
Which unlicensed users should be licensed first? How do Teams-heavy, Outlook-heavy, and Office-heavy users rank under different weighting profiles?

**M365 Usage Activity & Trends:**
How are users distributed across engagement segments? Is engagement trending up or down week over week?

</details>

---

<details>
<summary><strong style="font-size:1.5em;">🖥️ Report Pages Overview</strong></summary>

<br>

The dashboard includes **7 interactive report pages**. A detailed analysis guide will be available separately.

---

### 1. M365 Executive Summary

Your landing page — headline KPIs (total users, active Copilot users, departments, aggregate activity) and an app activity breakdown chart showing which M365 workloads are driving the most engagement.

[![M365 Executive Summary](images/1%20-%20M365%20Executive%20Summary.png)](images/1%20-%20M365%20Executive%20Summary.png)
*Click image to enlarge*

---

### 2. M365 Usage Trends

Week-over-week trend lines across Teams, Outlook, Word, Excel, and PowerPoint. A tier distribution matrix shows how users are spread across activity tiers per application, and a donut chart highlights proportional activity shares for the most recent week. Filter by date range, app, or department.

[![M365 Usage Trends](images/2%20-%20M365%20Usage%20Trends.png)](images/2%20-%20M365%20Usage%20Trends.png)
*Click image to enlarge*

---

### 3. Copilot License Recommendations

Every user ranked by a weighted composite LP (Licensing Priority) score blending their activity across M365 apps. Adjust weighting sliders or select preset profiles to model different licensing scenarios. The ranked table is ready for export and stakeholder review.

[![Copilot License Recommendations](images/3%20-%20Copilot%20License%20Recommendations.png)](images/3%20-%20Copilot%20License%20Recommendations.png)
*Click image to enlarge*

---

### 4. Copilot Enablement Strategy

A 2×2 quadrant model classifying every user into four segments: **Enablement Targets** (high M365 / low Copilot — best training candidates), **Champions** (high on both — peer advocates), **AI-First** (low M365 / high Copilot), and **Low Engagement** (low on both). Filter by app or quadrant to drill into individual users.

[![Copilot Enablement Strategy](images/4%20-%20Copilot%20Enablement%20Strategy.png)](images/4%20-%20Copilot%20Enablement%20Strategy.png)
*Click image to enlarge*

---

### 5. Copilot Enablement Strategy Tiers

A 5-level priority view — **Critical**, **High**, **Medium**, **Promoter**, **Low** — based on the gap between M365 activity and Copilot adoption per app. Users with the largest gap are the top priority for training investment. Includes per-user drill-down with recommended next actions.

[![Copilot Enablement Strategy Tiers](images/5%20-%20Copilot%20Enablement%20Strategy%20Tiers.png)](images/5%20-%20Copilot%20Enablement%20Strategy%20Tiers.png)
*Click image to enlarge*

---

### 6. M365 Usage Activity

Users classified into engagement segments by active days per week — Daily, Frequent, Moderate, Infrequent, and Inactive — cross-referenced with Copilot license status. A departmental comparison chart highlights organizational hotspots and cold spots across M365 workloads.

[![M365 Usage Activity](images/6%20-%20M365%20Usage%20Activity.png)](images/6%20-%20M365%20Usage%20Activity.png)
*Click image to enlarge*

---

### 7. Glossary & Metric Definitions

In-report reference covering every metric, quadrant, tier label, and scoring methodology used across the dashboard. Includes a visual heatmap matrix and complete definitions so the report can be shared broadly without external documentation.

[![Glossary and Metric Definitions](images/7%20-%20Glossary%20and%20Metric%20Definitions.png)](images/7%20-%20Glossary%20and%20Metric%20Definitions.png)
*Click image to enlarge*

</details>

---

<details>
<summary><strong style="font-size:1.5em;">📋 Instructions</strong></summary>

<br>

### Step 1. Export Your M365 Unified Audit Log Data (Required)

You need a CSV export of your organization's Unified Audit Log containing M365 usage activity. There are two ways to get this data:

#### Method 1: Manual Export from Microsoft Purview

Go to the [Microsoft Purview Portal](https://purview.microsoft.com/) and export your Unified Audit Log as a CSV.

<details>
<summary><strong>Detailed step-by-step guide</strong></summary>

<br>

### Quick Reference:

1. **Navigate to Audit Search**
   - Go to [https://purview.microsoft.com/](https://purview.microsoft.com/) (or [https://compliance.microsoft.com/](https://compliance.microsoft.com/))
   - In the left sidebar, select the **Audit** solution
   - Select the **Search** tab

2. **Configure Your Search**
   - **Date range**: Set a minimum of 90 days; 6–12 months recommended for meaningful tier ranking
   - **Activities**: Leave blank to capture all activity types (recommended), or filter to specific M365 usage workloads (Teams, Exchange, SharePoint, OneDrive, Office apps)
   - **Users**: Leave blank to capture the full tenant, or scope to a specific group
   - Click **Search** and wait for the job to complete

3. **Export the Results**
   - Once the search status shows **Completed**, click on the job name
   - Click **Export results** → **Download all results**
   - The exported CSV contains the standard Purview audit log columns: `CreationDate`, `UserIds`, `Operations`, and `AuditData`
   - Do **not** pre-process or re-save the file — column order and formatting must be preserved

4. **Rename and Place the File**
   - Rename the exported file or note its exact file path
   - You will enter this path as the `PurviewData` parameter in Power BI Desktop

> ⚠️ **Large tenants:** Purview audit log CSV exports are capped at **50,000 rows** per search job for Audit (Standard) or **100,000 rows** for Audit (Premium). For full coverage, run multiple searches with different date ranges and combine the CSVs, or use **Method 2** below which handles pagination automatically via the Microsoft Graph API.

</details>

---

#### Method 2: PAX Purview Audit Log Processor Script (Recommended)

The [PAX Purview Audit Log Processor](https://github.com/microsoft/PAX) is an open-source Microsoft PowerShell script that automates Purview unified audit log retrieval via the Microsoft Graph API. It handles pagination, rate limiting, and retry logic automatically. When run with the `-IncludeM365Usage` switch, it pulls all the operation types this dashboard needs to properly analyze your M365 usage data.

**Why choose this over manual exports?**

- **Can run on a schedule** — Set it up once, then automate it to run daily, weekly, or monthly so your dashboard data stays current without any manual effort
- **Runs completely unattended** — No need to sit at the keyboard; the script handles everything on its own from start to finish
- **No record limits** — Manual Purview exports are capped at 50K–100K rows per export; the script retrieves all available records, no matter how large your tenant
- **Automatically resumes if interrupted** — If a data pull is paused or fails for any reason, the script picks up right where it left off with no lost progress
- **Pulls exactly the right data** — Targets only the audit log activity types this dashboard requires, so your export is ready to use immediately
  <br>*When run with the recommended settings shown in the step-by-step guide below.*

<details>
<summary><strong>Detailed step-by-step guide</strong></summary>

<br>

1. **Download the script** from the [PAX GitHub repository](https://github.com/microsoft/PAX) — download the latest version of the `PAX_Purview_Audit_Log_Processor` script

2. **Run the script with the `-IncludeM365Usage` switch** to include all M365 usage activity types:

   **Interactive web login (easiest — no app registration required):**
   ```powershell
   .\PAX_Purview_Audit_Log_Processor.ps1 -IncludeM365Usage -StartDate "2025-01-01" -EndDate "2025-06-30"
   ```

   **App Registration (for scheduled / unattended runs):**
   ```powershell
   .\PAX_Purview_Audit_Log_Processor.ps1 -ClientId "<app-id>" -TenantId "<tenant-id>" -ClientSecret "<secret>" -IncludeM365Usage -StartDate "2025-01-01" -EndDate "2025-06-30"
   ```

3. **Locate the output** — The script outputs a CSV file to the current directory upon completion. The file path is displayed in the console output. Use this file path as the `PurviewData` parameter in Power BI Desktop.
> 💡 **The PAX script can also export Entra user and licensing data** needed by this dashboard — either alongside the Purview audit data in a single run or on its own. See **Option B** in the [Export Entra User Details](#step-2-export-entra-user-details) section below for details.
> � See the [PAX repository](https://github.com/microsoft/PAX) for full documentation, authentication setup guides, and advanced options.

</details>

---

### Step 2. Export Entra User Details

The report joins audit log data to Entra ID user attributes (department, license status, etc.) for filtering and Copilot license analysis. Without this data, department-level breakdowns and license-based insights will not be available.

<details>
<summary><strong>Detailed step-by-step guide</strong></summary>

<br>

#### Option A: Export from Entra Admin Center (Manual)

1. **Navigate to Microsoft Entra Admin Center**
   - Go to [https://entra.microsoft.com/](https://entra.microsoft.com/)
   - In the left sidebar, go to **Users** → **All users**

2. **Export User List**
   - Click **Download users** (top toolbar)
   - Select **All users** and include at minimum:
     - `userPrincipalName`
     - `displayName`
     - `department`
     - `jobTitle`
     - `assignedLicenses` (Option B below will automatically add a ready-to-use `hasLicense` column for you)
   - Download as CSV

#### Option B: PAX Purview Audit Log Processor Script (Recommended)

The same [PAX script](https://github.com/microsoft/PAX) used to pull Purview audit log data in Step 1 can also export all the Entra user and licensing details this dashboard needs. You can pull both datasets in a single script run, or export Entra user data on its own.

**Why choose this over manual exports?**

- **Can run on a schedule** — Automate regular exports so your user and licensing data stays up to date without manual effort
- **Runs completely unattended** — No need to sign in to any admin portal or sit at the keyboard
- **Automatically resumes if interrupted** — If the export is paused or fails, the script picks up where it left off with no lost progress
- **Pulls exactly the right user attributes** — Exports only the fields this dashboard needs, including Copilot license status, so the output is ready to import immediately
  <br>*When run with the recommended settings shown below.*

**How to run it:**

You can combine the Entra user export with the Purview audit log pull from Step 1 in a single run, or export Entra data only.

**Interactive web login (easiest — no app registration required):**

*Combined run — pulls both Purview audit data and Entra user/licensing details at once (separate output files are generated for each):*
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -IncludeM365Usage -IncludeUserInfo -StartDate "2025-01-01" -EndDate "2025-06-30"
```

*Entra users and licensing only — skips the Purview audit data pull entirely:*
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -OnlyUserInfo
```

**App Registration (for scheduled / unattended runs):**

*Combined run — pulls both Purview audit data and Entra user/licensing details at once (separate output files are generated for each):*
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -ClientId "<app-id>" -TenantId "<tenant-id>" -ClientSecret "<secret>" -IncludeM365Usage -IncludeUserInfo -StartDate "2025-01-01" -EndDate "2025-06-30"
```

*Entra users and licensing only — skips the Purview audit data pull entirely:*
```powershell
.\PAX_Purview_Audit_Log_Processor.ps1 -ClientId "<app-id>" -TenantId "<tenant-id>" -ClientSecret "<secret>" -OnlyUserInfo
```

> ℹ️ The `-OnlyUserInfo` switch does not require `-StartDate` or `-EndDate` because Entra user data is not date-ranged — it always exports the current state of your directory.

**Locate the output** — The Entra user and licensing data is saved to its own CSV file, separate from the Purview audit data CSV. The file path is displayed in the console output. Use this file path as the `EntraUsers` parameter in Power BI Desktop.

> � The exported CSV automatically includes a `hasLicense` column that checks each user's assigned licenses and flags whether they have a Microsoft 365 Copilot license. This means the dashboard can identify Copilot-licensed users right away — no extra steps needed on your end.

> �📖 See the [PAX repository](https://github.com/microsoft/PAX) for full documentation, authentication setup guides, and advanced options.

---

#### Option C: Export via PowerShell (Manual)

Run the following to produce the format the report expects. This creates the required `hasLicense` column in the output, but leaves it blank for all users by default:

   ```powershell
   Connect-MgGraph -Scopes "User.Read.All"
   Get-MgUser -All -Property userPrincipalName,displayName,department,jobTitle,assignedLicenses |
     Select-Object userPrincipalName, displayName, department, jobTitle,
       @{N='hasLicense'; E={ $null }} |
     Export-Csv -Path ".\EntraUsers.csv" -NoTypeInformation
   ```

> The export above includes the `hasLicense` column the dashboard expects, but the values will be empty. The dashboard will still load and function — you just won't see Copilot license–based insights until you populate that column.

**Want to populate the `hasLicense` column?** You'll need to look up your tenant's Microsoft 365 Copilot license SKU and check each user's assigned licenses against it. Here's how:

1. **Find your Copilot SKU ID** — Run this command while still connected to Microsoft Graph:
   ```powershell
   Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -like '*Copilot*' } |
     Select-Object SkuPartNumber, SkuId
   ```
   Copy the `SkuId` value from the output — you'll use it in the next step.

2. **Re-run the export with Copilot license detection** — Replace `<copilot-sku-id>` with the value you copied above:
   ```powershell
   $copilotSku = "<copilot-sku-id>"
   Get-MgUser -All -Property userPrincipalName,displayName,department,jobTitle,assignedLicenses |
     Select-Object userPrincipalName, displayName, department, jobTitle,
       @{N='hasLicense'; E={ ($_.AssignedLicenses.SkuId -contains $copilotSku) }} |
     Export-Csv -Path ".\EntraUsers.csv" -NoTypeInformation
   ```
   This sets `hasLicense` to `True` for users with a Copilot license and `False` for everyone else, giving the dashboard everything it needs for license-based analysis.

> 💡 **Option B (PAX script) handles all of this automatically** — no SKU lookup or extra steps required.

3. **Note the file path** — you will enter it as the `EntraUsers` parameter in Power BI Desktop.

> ℹ️ **Refresh cadence:** Re-export your Entra user data whenever there are significant changes to your user directory (new hires, departures, license changes, department reorganizations). For ongoing monitoring, a monthly re-export is recommended.

</details>

---

### Step 3. Open the Report in Power BI Desktop and Set Parameters

<details>
<summary><strong>Detailed step-by-step guide</strong></summary>

<br>

1. **Open the PBIP file**
   - Open **Power BI Desktop**
   - Go to **File** → **Open report** → **Browse** → select `M365.pbip` from this folder

2. **Set the Data Source Parameters**
   - Go to **Home** → **Transform data** → **Edit parameters**
   - Update the following parameters:
     | Parameter | Value |
     |---|---|
     | `PurviewData` | Full path to your Purview Unified Audit Log CSV (from Step 1) |
     | `EntraUsers` | Full path to your Entra user details CSV (from Step 2) |
   - Click **OK** → **Apply changes**

3. **Refresh the Report**
   - Click **Refresh** on the Home ribbon
   - On first load with a large audit log, allow several minutes for processing
   - Watch for any refresh errors in the **Transform data** pane

</details>

---

## Next Steps

<details>
<summary><strong>Validation & Troubleshooting</strong></summary>

**Checklist for success:**
- No errors on load
- Executive Summary visuals populate with user counts and tier distributions
- Copilot Enablement Strategy page shows users distributed across quadrants
- LP table on Copilot License Recommendations page shows ranked users with composite scores

**Common Mistakes & Fixes**

| Symptom | Cause | Fix |
|---|---|---|
| Blank matrix, no rows | Hardcoded slicer default filtering to a value that doesn't exist in tenant data | Open `visual.json` for the slicer and remove the `filter` property from `objects.general[0].properties` |
| Cards showing 0 | EntraData filtered to zero rows by a slicer default | Same fix as above — remove hardcoded department or group filter |
| LP table blank | Row field set to `M365Usage.UserId` instead of `EntraData.userPrincipalName` | Change the row field in the visual's `queryState.Rows` to `EntraData.userPrincipalName` |
| LP table blank after row field fix | Visual filter using the wrong filter measure for the page type | Replace `App Selection.Show User Based on App Tier Filter` with `LP Measures.LP Action Filter Match` |
| All users showing "Developing" | No M365 activity in the date range of the export | Widen the Purview export date range to at least 90 days |
| Tier ranking incorrect | Zero-activity users included in percentile calculation | Confirm `PERCENTILEX` filters to `[Measure] > 0` before ranking |
| Refresh type error on JSON column | Some audit event types have mixed integer/string fields | Use `try Int64.From(...) otherwise null` in the M Code extraction step |
| Report fails to open (relationship error) | Removed columns that had auto-date relationships | Remove the orphaned relationship blocks and LocalDateTable refs from `relationships.tmdl` and `model.tmdl` |
| Slicer shows blank/garbage first item | Header row not removed after `Binary.Decompress` step | Add `Table.Skip(#"Renamed Columns", 1)` after renaming columns |

See [`docs/M365_MCode_Troubleshooting.md`](docs/M365_MCode_Troubleshooting.md) for detailed diagnosis flowcharts and root cause write-ups on all known issues.

</details>

<details>
<summary><strong>Publish / Distribute</strong></summary>

- Save your PBIP file after setup.
- Publish to a Power BI workspace via **Home** → **Publish** in Power BI Desktop.
- After publishing, navigate to your workspace and open the **Semantic Model** settings.
- Configure the **Data source credentials** to point to the CSV file locations (local path or SharePoint/OneDrive URL).
- For scheduled refresh, the Purview CSV must be accessible from the gateway or cloud location at refresh time — store the file in SharePoint or OneDrive for cloud refresh compatibility.

</details>

<details>
<summary><strong>Interpretation & Action Planning</strong></summary>

For a visual walkthrough of each report page — including what it shows and how to read it — see the [Report Pages Overview](#report-pages-overview) section above.

Use the report pages in this order to tell a complete Copilot readiness story:

1. **Executive Summary** — Overall tenant snapshot: total users, Copilot users, department count, and aggregate activity by workload
2. **M365 Usage Trends** — How engagement is trending week over week across each app, with tier distribution per application
3. **Copilot License Recommendations** — Ranked list of top candidates for net-new Copilot licenses, with adjustable weighting profiles
4. **Copilot Enablement Strategy** — 2×2 quadrant view: identify Enablement Targets, Champions, AI-First users, and Low Engagement cohorts
5. **Copilot Enablement Strategy Tiers** — Priority-level view: Critical, High, Medium, Promoter, Low — focus training on the largest M365-to-Copilot gaps
6. **M365 Usage Activity** — Baseline engagement segments and departmental comparisons across all workloads
7. **Glossary & Definitions** — Tier definitions, quadrant thresholds, action labels, and composite scoring methodology

</details>

<details>
<summary><strong>Monitor with Automatic Refresh</strong></summary>

- After publishing to Power BI Service, configure the Semantic Model for scheduled refresh.
- Navigate to [Power BI Web](https://app.powerbi.com/) and locate the published Semantic Model.
- Hover over the Semantic Model → click the **Schedule refresh** icon.
- Configure refresh frequency (weekly recommended to stay in sync with Purview export cadence).
- For ongoing monitoring, re-run your Purview Audit Log export regularly (monthly or on a rolling schedule), overwrite the previous CSV file at the same path, and the report will pick up the new data on next refresh.
- If using the PAX script, consider scheduling it as a recurring task (e.g., Windows Task Scheduler or Azure Automation) to automate the data pipeline end to end.
- Track changes in tier distribution over time — users moving from Bottom 50% into Top 25% is a leading indicator of Copilot adoption success.

</details>

</details>

---

<details>
<summary><strong style="font-size:1.5em;">🤓 Nerd Corner</strong></summary>

<br>

If you're into automation and allergic to manual decks — try this:

👉 https://github.com/shailendrahegde/pbi-to-exec-deck

It turns raw Power BI outputs into **exec-ready PPTs** with insights pre-baked.
All you do: verify, tweak, ship.

Also see [`docs/M365_MCode_Build_Strategy.md`](docs/M365_MCode_Build_Strategy.md) for a full technical breakdown of the M Code pipeline, DAX architecture, and scaling checklist for deploying to new tenants.

</details>

---

<details>
<summary><strong style="font-size:1.5em;">💬 Feedback</strong></summary>

<br>

Managed and released by the Microsoft Copilot Growth ROI Advisory Team. Please reach out to [copilot-roi-advisory-team-gh@microsoft.com](mailto:copilot-roi-advisory-team-gh@microsoft.com) with any feedback.

</details>

---

<details>
<summary><strong style="font-size:1.5em;">🔔 Stay Updated</strong></summary>

<br>

- ⭐ **Star this repository** to receive notifications about new template versions
- 👀 **Watch** for updates and announcements
- 🔄 Check back regularly for new features and improvements

</details>

---
