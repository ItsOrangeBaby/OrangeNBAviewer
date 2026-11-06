# Orange NBA Viewer — Reusable Daily Slate App

**This app is not restricted to November 6 or any other specific date.** Use the same `app.py` every NBA slate. Upload new files and select the correct game date; no code edits or GitHub commits are needed for daily slate changes.

## Install once

1. Add `app.py` and `requirements.txt` from this package to the root of your `Orange-NBA-Viewer` GitHub repository.
2. Commit the files and allow Streamlit Community Cloud to redeploy.
3. Leave the app deployed. Do not upload your daily NBA workbooks or CSV exports to GitHub: upload them directly in the app for each slate.

Dependencies in `requirements.txt` are Streamlit and pandas. The app reads `.xlsx` directly without requiring Excel or a fixed local workbook.

## Every NBA slate — repeat these steps

1. **Prepare today's NBA statistics workbook** as an `.xlsx` with exactly two sheets: `Player Stats` and `Team Stats`, using the existing Orange NBA column layout. It can contain any number of slate teams and players; it is not limited to 9 games, 18 teams, or 260 players.
2. **Open Orange NBA Viewer** and upload the new workbook in **NBA master workbook (.xlsx)**.
3. **Check `Slate date (U.S. Eastern)`** in the Game slate section. The app infers a date from the workbook filename if possible; otherwise it defaults to the current U.S. Eastern date. You can change the date manually at any time.
4. **Use `Automatic daily schedule`** to fetch that date's NBA games and identify opposing defenses. The app uses the schedule feed only for matchups; all NBA statistics are taken from your workbook. If the feed is unavailable or a team is missing, select **Enter / edit matchups** and supply the pairings yourself.
5. **Upload that slate's Roto Projections CSV.** This file is required for player eligibility: only players listed in the selected projection pool will be shown in Player Stats and Player Matchup rankings. If available, choose MAIN, LATE, or All pools.
6. **Upload that slate's Optimized Lineups CSV** to add lineup appearances and 0–15 Roto points. If no optimizer CSV is provided, Roto remains blank rather than being treated as a zero.
7. **Review** Team Highlights, Team Stats, Matchup Score, Player Stats, and Player Matchup Adjustments & Rankings. Choose the games to include in the schedule selector. On the next slate, repeat from step 1; keep the same Python script.

### Filename conventions and date safeguards

The app recognizes dates in names such as:

- `NBA_Stats_Nov_07_2026.xlsx`
- `NBA_Stats_2026-11-07.xlsx`
- `nba_lineups_20261107.csv`
- `nba_projections_11-07-2026.csv`

A file without a recognizable date can still be uploaded. The app cannot automatically determine its date from a filename such as `3350865(1).csv`, so confirm it belongs to your selected slate. If two uploaded filenames contain different detectable dates, the app warns you. This is a warning, not an automatic override of your chosen date.

The slate date is based on the U.S. Eastern NBA calendar, even if your local time zone differs.

## What stays the same every day

- **Team Stats:** team outlier highlights, Boxes 1–4 together on one page, and Box 5 reserved for future outlier statistics and rules.
- **Matchup Score:** Paint, Shooting, Scoring, and Fantasy, using all 24 established team scoring statistics, with partial coverage indicated for missing source values.
- **Player Stats:** nine groups of six original statistics, filtered to the selected projections roster. Projection salary, floor, ceiling, FPTS/$, Smash, and lineup counts are shown when available.
- **Player Matchup Adjustments:** uses each player's actual selected-slate opponent and your existing defensive threshold multipliers. Multiple relevant multipliers are multiplied together, as requested.
- **Overall Player Score:** sums all valid adjusted Paint, Shooting, Attacking, and Fantasy statistics, plus available Roto lineup points. The existing 15 Roto point tiers are unchanged.
- **Player-name aliases:** approved alias matching remains active and team-aware.

## Blank statistics and projection-only players

- Blank or invalid individual NBA statistics remain blank in detail tables; **they are not converted to zero**.
- Other valid adjusted statistics continue contributing to the player’s overall score, which can be labelled **Partial**.
- A player who is on the projections CSV but absent from the NBA workbook stays eligible. Their NBA fields remain blank; if there are valid lineup appearances, they can get a **Roto only** score.
- A player **not on the selected projections CSV** is excluded from the player sections, regardless of their workbook statistics or optimizer appearances.
- A real zero optimizer appearance is used only when valid optimizer entries match the selected projection pool. Without a matching optimizer file, the Roto score remains blank.

## If the results look incomplete

- **Wrong teams/opponents:** check the `Slate date`, selected games, and that the workbook teams correspond to the selected date. Use the manual matchup editor if needed.
- **Player missing:** confirm the player is in the chosen projections pool, and check the name-matching warnings.
- **Partial player score:** source statistics are missing or invalid; all other available values still count.
- **No Roto points:** check that the optimized lineups CSV belongs to the same slate/pool as the projections CSV.
- **Box 5 empty:** expected until you specify the six advanced team statistics and their outlier rules.

## Validation notes

The reusable script was syntax-checked and its daily date parser tested with multiple filename/date formats. Its player inclusion, missing-stat handling, and category scoring were regression-tested with a sample workbook and CSV exports. **Those files are only test examples—not built-in requirements.**

Real-time schedule access and the live deployed Streamlit interface have not been end-to-end tested in this environment. If the schedule provider is unavailable, the manual matchup editor is the fallback.
