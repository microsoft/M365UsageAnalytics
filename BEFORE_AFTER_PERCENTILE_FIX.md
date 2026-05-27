# Performance fix spec — Enablement Strategy + License Allocation pages

**Customer symptom:** 70K-user tenant, 1 week of data → both pages fail to render.
**Root cause:** 5 measures recompute percentiles at query time using `PERCENTILEX` / `CALCULATETABLE(ALL(...))` patterns that scale O(N²). The Python processor already pre-computes equivalent values into UserStats and these are already exposed as columns on `M365Usage`, but the measures don't use them.

---

## Pre-computed columns already available on `M365Usage`

Confirmed present in [`M365Usage.tmdl`](DEMO/M365 Dashboard May 21st V3 with Optimizer.SemanticModel/definition/tables/M365Usage.tmdl) (joined in via Power Query from `UserStats.csv`):

| Column | Source (UserStats) | Semantic |
|---|---|---|
| `Copilot Usage Rank Column` | `CopilotUsageRankColumn` | 0–1 normalized rank of user's total Copilot events vs. all users |
| `M365 Usage Rank Column` | `M365UsageRankColumn` | 0–1 normalized rank of user's total M365 events vs. all users |
| `M365 Tier Column` | `M365TierColumn` | Pre-classified: `1` (Top 10%) / `2` (10–25%) / `3` (25–50%) / `4` (Bottom 50%) |
| `Copilot Tier Column` | `CopilotTierColumn` | Same buckets, Copilot-based |
| `Priority Scatter Column` | `PriorityScatterColumn` | Composite priority bucket |
| `Is Copilot User` | `IsCopilotUser` | Has any Copilot event in window |
| `Session Cohort` | `SessionCohort` | Bucketed session count |

**Important semantic note:** all four rank/tier columns are computed at **all-apps total** granularity in the Python script (per-user totals, not per-(user, app)). They are **not** parameterized by the App Selection slicer. See "Semantic trade-off" at the end.

---

## Measure 1 — `CE App Percentile`

**File:** [`Copilot Measures.tmdl` line 219](DEMO/M365 Dashboard May 21st V3 with Optimizer.SemanticModel/definition/tables/Copilot%20Measures.tmdl#L219)

### BEFORE — O(N²) recompute

```dax
measure 'CE App Percentile' =
    VAR UserScore = [Selected App M365 Activity]
    VAR AllScores =
        CALCULATETABLE(
            ADDCOLUMNS(
                FILTER(
                    ALL(EntraUsers[userPrincipalName]),
                    [Selected App M365 Activity] > 0
                ),
                "@s", [Selected App M365 Activity]
            ),
            ALL(EntraUsers)
        )
    VAR Below = COUNTROWS(FILTER(AllScores, [@s] <= UserScore))
    VAR Total = COUNTROWS(AllScores)
    RETURN
        IF(UserScore > 0 && Total > 0,
           ROUND(DIVIDE(Below, Total) * 100, 0),
           BLANK())
```

For each visible row: scans all 70K users, recomputes `Selected App M365 Activity` for each → ~70K × 70K measure evaluations per page render.

### AFTER — scalar lookup of pre-computed rank

```dax
measure 'CE App Percentile' =
    VAR Rank01 =
        CALCULATE(
            MAX('M365Usage'[M365 Usage Rank Column]),
            REMOVEFILTERS('M365Usage'[Operation]),
            REMOVEFILTERS('M365Usage'[AppHost]),
            REMOVEFILTERS('M365Usage'[Workload])
        )
    RETURN
        IF(NOT ISBLANK(Rank01),
           ROUND(Rank01 * 100, 0),
           BLANK())
```

Reads one pre-computed column. **Expected speedup: 50–500× on 70K users.**
`MAX` instead of `SELECTEDVALUE` because the same `Rank01` value repeats across every row for a given user; `MAX` collapses ties safely.

---

## Measure 2 — `CE Copilot Percentile`

**File:** [`Copilot Measures.tmdl` line 247](DEMO/M365 Dashboard May 21st V3 with Optimizer.SemanticModel/definition/tables/Copilot%20Measures.tmdl#L247)

### BEFORE — O(N²) recompute (worst offender — nested `CALCULATE` inside `ADDCOLUMNS`)

```dax
measure 'CE Copilot Percentile' =
    VAR UserCopilot =
        CALCULATE(SUM('M365Usage'[EventCount]),
                  M365Usage[Workload] = "Copilot"
                      || CONTAINSSTRING(M365Usage[Operation], "CopilotInteraction"))
    VAR AllScores =
        CALCULATETABLE(
            ADDCOLUMNS(
                FILTER(
                    ALL(EntraUsers[userPrincipalName]),
                    CALCULATE(SUM('M365Usage'[EventCount]),
                              M365Usage[Workload] = "Copilot"
                                  || CONTAINSSTRING(M365Usage[Operation], "CopilotInteraction")) > 0
                ),
                "@s", CALCULATE(SUM('M365Usage'[EventCount]),
                                M365Usage[Workload] = "Copilot"
                                    || CONTAINSSTRING(M365Usage[Operation], "CopilotInteraction"))
            ),
            ALL(EntraUsers)
        )
    VAR Below = COUNTROWS(FILTER(AllScores, [@s] <= UserCopilot))
    VAR Total = COUNTROWS(AllScores)
    RETURN
        IF(UserCopilot > 0 && Total > 0,
           ROUND(DIVIDE(Below, Total) * 100, 0),
           BLANK())
```

### AFTER

```dax
measure 'CE Copilot Percentile' =
    VAR Rank01 =
        CALCULATE(
            MAX('M365Usage'[Copilot Usage Rank Column]),
            REMOVEFILTERS('M365Usage'[Operation]),
            REMOVEFILTERS('M365Usage'[AppHost]),
            REMOVEFILTERS('M365Usage'[Workload])
        )
    RETURN
        IF(NOT ISBLANK(Rank01),
           ROUND(Rank01 * 100, 0),
           BLANK())
```

**Expected speedup: 50–500×.**

---

## Measure 3 — `CE Enablement Targets Count`

**File:** [`Copilot Measures.tmdl` line ~194](DEMO/M365 Dashboard May 21st V3 with Optimizer.SemanticModel/definition/tables/Copilot%20Measures.tmdl)

### BEFORE — iterates all users + calls both percentile measures per row

```dax
measure 'CE Enablement Targets Count' =
    VAR SelectedQuad = SELECTEDVALUE('CE Quadrant Selection'[QuadrantName], "Enablement Targets")
    VAR UsersWithScores =
        ADDCOLUMNS(
            ALLSELECTED(EntraUsers[userPrincipalName]),
            "@AppPct", [CE App Percentile],
            "@CopilotPct", [CE Copilot Percentile]
        )
    RETURN
        COUNTROWS(
            FILTER(UsersWithScores,
                NOT(ISBLANK([@AppPct])) && NOT(ISBLANK([@CopilotPct])) &&
                SWITCH(SelectedQuad,
                    "Enablement Targets", [@AppPct] >= 50 && [@CopilotPct] < 50,
                    "Champions",          [@AppPct] >= 50 && [@CopilotPct] >= 50,
                    "AI-First",           [@AppPct] < 50 && [@CopilotPct] >= 50,
                    "Low Engagement",     [@AppPct] < 50 && [@CopilotPct] < 50,
                    TRUE())))
```

After the BEFORE fix to measures 1 and 2, this becomes ~70K × 2 column lookups — already fast. But we can make it a pure SUMMARIZE over the precomputed columns.

### AFTER

```dax
measure 'CE Enablement Targets Count' =
    VAR SelectedQuad = SELECTEDVALUE('CE Quadrant Selection'[QuadrantName], "Enablement Targets")
    VAR UserRanks =
        SUMMARIZE(
            'M365Usage',
            'M365Usage'[UserId],
            "@AppPct",     MAX('M365Usage'[M365 Usage Rank Column])     * 100,
            "@CopilotPct", MAX('M365Usage'[Copilot Usage Rank Column])  * 100
        )
    RETURN
        COUNTROWS(
            FILTER(UserRanks,
                NOT ISBLANK([@AppPct]) && NOT ISBLANK([@CopilotPct]) &&
                SWITCH(SelectedQuad,
                    "Enablement Targets", [@AppPct] >= 50 && [@CopilotPct] < 50,
                    "Champions",          [@AppPct] >= 50 && [@CopilotPct] >= 50,
                    "AI-First",           [@AppPct] < 50 && [@CopilotPct] >= 50,
                    "Low Engagement",     [@AppPct] < 50 && [@CopilotPct] < 50,
                    TRUE())))
```

Single scan of the user dimension, no nested measure calls.

---

## Measures 4–8 — `Word/Excel/PowerPoint/Outlook/M365 Word,Excel,PPT Tier V2`

**File:** [`M365 Total Activity Measures LL.tmdl`](DEMO/M365 Dashboard May 21st V3 with Optimizer.SemanticModel/definition/tables/M365%20Total%20Activity%20Measures%20LL.tmdl)

### BEFORE — same anti-pattern, 3 PERCENTILEX calls per row

```dax
measure 'Word Tier V2' =
    VAR RawScore = [Word Activity Count v2]
    VAR P90 = PERCENTILEX.INC(FILTER(ALL('EntraUsers'), …), [Word Activity Count v2], 0.9)
    VAR P75 = PERCENTILEX.INC(FILTER(ALL('EntraUsers'), …), [Word Activity Count v2], 0.75)
    VAR P50 = PERCENTILEX.INC(FILTER(ALL('EntraUsers'), …), [Word Activity Count v2], 0.5)
    RETURN SWITCH(TRUE(), RawScore >= P90, 1, RawScore >= P75, 2, RawScore >= P50, 3, 4)
```

### AFTER (Option A — uses existing precomputed all-apps tier)

```dax
measure 'Word Tier V2' =
    CALCULATE(
        MAX('M365Usage'[M365 Tier Column]),
        REMOVEFILTERS('M365Usage'[Operation]),
        REMOVEFILTERS('M365Usage'[AppHost]),
        REMOVEFILTERS('M365Usage'[Workload])
    )
```

**Semantic change:** "Tier V2" was previously per-app; this collapses it to all-apps M365 tier. See trade-off section.

### AFTER (Option B — pre-compute per-app tiers in Python, no semantic change)

Requires processor update to add 5 new columns to UserStats:
`WordTierColumn`, `ExcelTierColumn`, `PowerPointTierColumn`, `OutlookTierColumn`, `TeamsTierColumn`

Then:
```dax
measure 'Word Tier V2' =
    CALCULATE(
        MAX('M365Usage'[Word Tier Column]),
        REMOVEFILTERS('M365Usage'[Operation]),
        REMOVEFILTERS('M365Usage'[AppHost]),
        REMOVEFILTERS('M365Usage'[Workload])
    )
```

---

## Semantic trade-off — please confirm before I implement

The current `CE App Percentile` and `*Tier V2` measures vary their answer based on the **App Selection slicer** (e.g. if user picks "Word", percentile is computed only over Word activity). The pre-computed `M365 Usage Rank Column` from UserStats is **all-apps total**, not per-app.

**Two paths:**

| | Option A — Use existing precomputed columns | Option B — Extend Python script |
|---|---|---|
| Semantic change | App-slicer no longer changes percentile (always shows all-apps M365 rank) | None — preserves current behavior |
| DAX changes | 8 measures rewritten as shown above | Same 8 measures + reference new per-app columns |
| Python script changes | None | Add 5 per-app rank/tier columns to UserStats |
| PBIT refresh impact | Immediate (no re-export needed) | Customer must re-run processor on new CSV |
| Customer perf gain | 50–500× on the 2 broken pages | Same |
| Risk | Story for "Enablement Strategy" page becomes "M365 power users not using Copilot" (still the right insight) | Lower — pure perf fix |

**My recommendation:** **Option A** for the immediate hotfix. The Enablement Strategy quadrant story is fundamentally about "who's a heavy M365 user but a light Copilot user" — collapsing to all-apps rank actually makes the quadrant story clearer (the per-app variant was already inconsistent with the all-apps Copilot rank on the y-axis). Ship Option B in a follow-up release if a customer specifically requests per-app drilling.

---

## Implementation plan (once you sign off)

1. Branch `perf/precomputed-percentiles` off `release/may-21-2026-clean`.
2. Edit the 2 tmdl files in `DEMO\M365 Dashboard May 21st V3 with Optimizer.SemanticModel\definition\tables\` (this is your working PBIP — the published `.pbit` will be regenerated from it).
3. Open the PBIP in Power BI Desktop, confirm the Enablement Strategy and License Allocation pages render and the numbers look right against your demo data.
4. Re-export the `.pbit`.
5. Replace `M365 Usage Dashboard - Power BI TEMPLATE - 21 May 2026.pbit` at the repo root.
6. Commit + push to your fork + open PR to `microsoft/main`.

No Python script changes needed for Option A.
