# Orange NBA Viewer — Reusable Daily Slate (No Player Multipliers)

The same Streamlit app works for **any NBA slate date**. This update **removes all opponent-based player multipliers** (previously 0.5× / 1.0× / 1.5×). NBA player scores now use the **original, unadjusted workbook statistics**, plus Roto points from your uploaded optimized-lineup CSV. Opposing teams remain visible for context and optional game filtering, but **their defensive tier does not change player statistics**.

## Install once

1. Replace `app.py` and `requirements.txt` at the root of your `Orange-NBA-Viewer` GitHub repository with the files in this package.
2. Commit the changes and allow Streamlit Community Cloud to redeploy.
3. Leave the app deployed. For each new slate, **upload new files in the app**, not to GitHub.

## Repeat for every slate

1. Upload the day's `.xlsx` workbook containing **exactly two sheets**: `Player Stats` and `Team Stats`, with the existing Orange NBA column layout. The number of games, players, and teams can change daily.
2. Confirm the **Slate date (U.S. Eastern)**. The app recognizes dates in workbook filenames such as `NBA_Stats_Nov_06_2026.xlsx` or `NBA_Stats_2026-11-07.xlsx`; you can always change the date manually.
3. Use the **Automatic daily schedule** to show and filter the NBA matchups. If the schedule is unavailable, you may enter opponents manually. **Player rankings still calculate even without opponent assignments**; opponent labels are simply blank.
4. Upload the day's **Roto Projections CSV**. Only players in the selected projection pool appear in Player Stats or Player Rankings. If the CSV includes MAIN and LATE pools, select the appropriate pool or All pools.
5. Upload the day's **Optimized Lineups CSV** to count player appearances and add the existing **0–15 Roto tier score**. Without an optimizer file, Roto remains blank rather than zero.
6. Review Team Stats, Matchup Score, Player Stats, and Player Rankings. Repeat these steps for the next slate without editing Python.

## What scores mean

**Player Rankings** has Paint, Shooting, Attacking, Fantasy, Roto, and Overall Leaderboard tabs. Its score is:

**Overall Score = Paint + Shooting + Attacking + Fantasy + Roto (when provided)**

- **Paint:** 2PM, 2PA, RA FGM, RA FGA, NRA FGM, NRA FGA.
- **Shooting:** 3PM, 3PA, C3PM, C3PA, MR FGM, MR FGA.
- **Attacking:** USG%, PPG, FGM, FGA, %OREB, FGM %UAST.
- **Fantasy:** STL, BLK, %DREB, %AST, DD2, TD3.
- **Roto:** Points awarded from lineup appearances using your original 15 scoring bands; 225 or more appearances earns 15 points.

The app uses original **per-game numeric values** for all scoring stats. **No matchup tier multiplies, discounts, or inflates player statistics.** The Overall Score is a custom index that sums raw percentages, counts, and points. It is **not** a projected NBA point total or standard fantasy-point projection. Some measures overlap, as they did in the original score design.

### Blank statistics and projection eligibility

- Blank, invalid, or conflicting player statistics **stay blank** in tables and CSV exports; they **do not become zero** or invalidate the whole player's score.
- Valid statistics still add to the category and overall totals. The resulting score is marked **Partial** when NBA scoring statistics are missing.
- A genuine source value of `0` stays zero and counts as a valid statistic.
- Players on the projections list but missing from the NBA workbook are **retained** with blank NBA fields and may receive a **Roto only** score when lineup data exists.
- Players **absent from the projections CSV** are excluded, even if they appear in the NBA workbook or optimizer output.
- FGM is derived from workbook `2PM + 3PM` when needed (subject to rounding). Player Corner 3PA is not available in the established two-sheet template and is omitted consistently from Overall Score.

## Unchanged team sections

- **Team Stats:** favorable and unfavorable flags, Boxes 1–4 on one page, Box 5 reserved for future outlier fields.
- **Matchup Score:** all **24 unique team defense statistics** scored across Paint, Shooting, Scoring, and Fantasy; category maximum 20 points, overall maximum 80 points.
- **Team favorable/unfavorable threshold colors and scoring remain unchanged.** Removing *player* multipliers does **not** remove the Team Matchup Score.
- Team-specific player-name aliases, projections fields (salary, floor, ceiling, FPTS/$, Smash), and daily schedule lookup remain in place.

## Files and validation

- `app.py` — reusable Streamlit app with **no player multipliers**.
- `requirements.txt` — required Python packages.

Validated with the supplied November 6 workbook and both supplied CSVs for MAIN, LATE, and All pools. The test confirmed all 238 projection-eligible players are retained in the All pools view, raw stat values are unchanged by opponent assignments, partial scores still work, and team scoring is unchanged. The Python syntax check passed. The live hosted Streamlit interface and live schedule request have not been tested in this environment.
