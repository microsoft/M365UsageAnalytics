# M365 Usage Dashboard - v2.5.0 Fast Version (PATCHED_v3.pbit)

## What's Inside

### Precomputed Columns (39 new columns in EntraUsers)
These are calculated once by the Python script and stored as static lookup values:

**Raw Activity Columns (18 total):**
- `TeamsRaw_L30`, `TeamsRaw_L60`, `TeamsRaw_Full` — Sum of Teams activity per user in each window
- `OutlookRaw_L30`, `OutlookRaw_L60`, `OutlookRaw_Full` — Sum of Outlook activity
- `WordRaw_L30`, `WordRaw_L60`, `WordRaw_Full` — Sum of Word activity  
- `ExcelRaw_L30`, `ExcelRaw_L60`, `ExcelRaw_Full` — Sum of Excel activity
- `PowerPointRaw_L30`, `PowerPointRaw_L60`, `PowerPointRaw_Full` — Sum of PowerPoint activity
- `CopilotChatRaw_L30`, `CopilotChatRaw_L60`, `CopilotChatRaw_Full` — Sum of Copilot Chat activity

**Ranking Columns (18 total):**
- `CERank_Teams_L30/L60/Full` through `CERank_PowerPoint_L30/L60/Full` — User's percentile rank (0-100) for each app/window
- `CERank_M365AllApps_L30/L60/Full` — User's percentile rank for all M365 apps combined

**Percentile Columns (3 total):**
- `CECopilotPercentile_L30/L60/Full` — User's Copilot adoption percentile (0-100)

### New RankWindow Table
Allows users to switch between L30 (Last 30 days), L60 (Last 60 days), and Full (entire range):
- Code: "L30", "L60", "Full"
- Window: Display label
- Sort: Sort order

### 7 New Fast Measures
All use precomputed columns instead of expensive SUMX operations:

**Copilot Measures:**
1. `CE App Percentile Fast` — Fast version of original CE App Percentile
2. `CE Copilot Percentile Fast` — Fast version of original CE Copilot Percentile

**LP (License Planner) Weighted Measures:**
3. `LP Teams Weighted Fast` — Uses `TeamsRaw_*` columns
4. `LP Outlook Weighted Fast` — Uses `OutlookRaw_*` columns
5. `LP Word Weighted Fast` — Uses `WordRaw_*` columns
6. `LP Excel Weighted Fast` — Uses `ExcelRaw_*` columns
7. `LP PowerPoint Weighted Fast` — Uses `PowerPointRaw_*` columns

All new Fast measures include a `SWITCH` statement that selects the right window (L30/L60/Full) based on `RankWindow[Code]`.

### Original Measures: UNCHANGED
All original measures (CE App Percentile, CE Copilot Percentile, LP Teams Weighted, etc.) remain exactly as they were. This ensures backward compatibility and lets you compare performance side-by-side.

## How to Test

### Step 1: Open the PBIT
- Open `M365 Usage Dashboard - Power BI TEMPLATE - 22 May 2026.PATCHED_v3.pbit` in Power BI Desktop
- Fill in all parameters (Rollup, SessionCohort, UserStats paths, exportUsers CSV path)
- Point to your test data (either `v2_3_0_test` or `v220_test/Purview Downloads`, whichever has all 4 files)

### Step 2: Check that pages load
- Navigate to Enablement Strategy and License Recommendations pages
- These should load without errors (no Broken_Filters)
- Visuals may show slow/blank initially — this is normal while the measures recalculate

### Step 3: (Optional) Test the Fast Measures
To see if the new Fast measures work:
1. In Power BI Desktop, open the Enablement Strategy table visual
2. Edit the visual and swap one column:
   - Remove `CE App Percentile` column
   - Add `CE App Percentile Fast` column
3. Watch the refresh time — it should be instant if using precomputed columns
4. Verify the data looks correct

Repeat for other Fast measures on License Recommendations page.

## What This Solves

- ✅ Adds 39 precomputed columns (data pre-calculated offline in Python)
- ✅ Adds RankWindow table (L30/L60/Full switching)
- ✅ Provides 7 new Fast measures that use lookups instead of expensive SUMX operations
- ✅ Preserves all original measures (no breaking changes)
- ✅ Reduces dynamic calculation load on Enablement Strategy and License Recommendations pages

## Performance Impact

Expected improvements:
- **Original measures**: Still slow (dynamic calculation) but available for testing
- **Fast measures**: Should load instantly on 70K+ users (precomputed lookups)
- If visuals using Fast measures still show Broken_Filters → the issue is not measure expressions but something else (data source, schema validation, etc.)

## Next Steps

1. **If this loads clean with no errors**: Switch all visuals to use the Fast measures and we're done.
2. **If Broken_Filters still appears**: We know the issue isn't the measure rewrites themselves, but something deeper (data-source binding, annotations, etc.). Can then try isolated tests.
3. **If load is still slow even with Fast measures**: The bottleneck is elsewhere (Power Query refresh, data import, etc.).
