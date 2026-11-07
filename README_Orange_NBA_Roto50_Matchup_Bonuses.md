# Orange NBA Viewer — Reusable Daily Slate: Roto 50 + Opponent Matchup Bonuses

This `app.py` works with daily NBA Excel workbooks and two optional CSV sources, regardless of filename. It is **not** tied to November 7 (or any other date).

## Install and use

1. Replace the `app.py` file in your GitHub `Orange-NBA-Viewer` repository with the attached updated `app.py`.
2. Keep the supplied `requirements.txt`, or retain your existing compatible dependencies.
3. Commit and allow Streamlit Community Cloud to redeploy.
4. Upload that day's **NBA stats `.xlsx`**, **Optimized Lineups `.csv`**, and **Roto Projections `.csv`** through the three always-visible upload boxes.
5. In the game-slate selector, choose the correct date and verify the NBA game pairings. You may manually enter or edit pairings if necessary.
6. Select a projection pool when offered. Only players in the selected projection pool appear in Player Stats and Player Rankings.

### Workbook and CSV requirements

- Excel: `Team Stats` and `Player Stats` sheets in the established layout. Supplemental matchup tables below the main grids are ignored.
- Roto projections: a player/name column, team column, and preferably `OPP`, `MINUTES`, `FPTS`, `FPTS/$`, `OPTO`, `SMASH` plus the other existing projection fields. Header aliases and manual mapping controls are supported.
- Lineup CSV: eight roster-slot columns, a player-per-lineup table, or a precomputed appearance count. Uploads are checked by their contents, not by filename.

## Player scoring (170 maximum)

The existing **base player ratings** are unaffected by team stats. For each category, two priority player metrics score 0/3/6 each and four supporting metrics score 0/1/2 each, by player rank among projection-eligible athletes in the selected game scope. Base category maximum: **20**.

Now add **half the opposing team's 0–20 matchup category score** to the player's matching base category. This is **addition**, not multiplication. No player multipliers are used.

| Player category | Base maximum | Opposing team category used | Bonus maximum | Total maximum |
| --- | ---: | --- | ---: | ---: |
| Paint | 20 | Paint Score | 10 | 30 |
| Shooting | 20 | Shooting Score | 10 | 30 |
| Attacking | 20 | Scoring Score (team name) | 10 | 30 |
| Fantasy | 20 | Fantasy Score | 10 | 30 |
| **Roto** | — | Six CSV-based components | — | **50** |
| **Overall** | | | | **170** |

Example: if the player's base Paint score is 15 and **their opponent's** Paint Matchup Score is 18, final Paint is `15 + (18 / 2) = 24 / 30`.

Matchup source scores always come from the **opposing team's** 24 team defensive statistics and existing thresholds. The app prioritizes the selected daily slate; if the schedule cannot assign an opponent, it can use the `OPP` column in the projections CSV, but only when that opposing team appears in the workbook. Unresolved matchup bonuses remain blank, not 0.

If player stats, a particular defensive matchup metric, or a projection field is missing or invalid, only that component is skipped. Other available points still add, the player is ranked, and partial coverage is indicated. A genuine reported numeric zero still scores zero.

## Roto — 50 points maximum

| Component | Maximum | Source |
| --- | ---: | --- |
| Lineup Points | 15 | Number of optimizer lineup appearances |
| Opto Points | 5 | `OPTO` |
| Smash Points | 5 | `SMASH` |
| PPD Points | 10 | `FPTS/$` |
| Projected FPTS Points | 10 | `FPTS` |
| **Projected Minutes Points** | **5** | **`MINUTES`** |
| **Total Roto** | **50** | |

Projected minutes are scored using inclusive minimum thresholds:

| Projected minutes | Points |
| --- | ---: |
| Below 24 | 0 |
| 24 to less than 28 | 1 |
| 28 to less than 32 | 2 |
| 32 to less than 35 | 3 |
| 35 to less than 37 | 4 |
| 37 and above | 5 |

All previously defined Roto thresholds remain unchanged:

- **Lineup appearances:** 0=0; 1–9=1; 10–24=2; 25–44=3; 45–64=4; 65–84=5; 85–99=6; 100–115=7; 116–130=8; 131–145=9; 146–160=10; 161–175=11; 176–194=12; 195–210=13; 211–224=14; 225+=15.
- **Opto and Smash:** minimum 10=1, 16=2, 21=3, 26=4, 31=5.
- **Price per Dollar:** minimum 2.5=1, 3=2, 3.5=3, 4=4, 4.4=5, 4.8=6, 5.2=7, 5.6=8, 5.8=9, 6=10.
- **Projected FPTS:** minimum 19=1, 24=2, 29=3, 34=4, 39=5, 44=6, 49=7, 54=8, 59=9, 60=10.

## Screens in the app

- **Team Stats:** original team boxes and positive/negative highlights. Team Box 5 remains reserved until you define its statistics.
- **Matchup Score:** Paint, Shooting, Scoring, Fantasy, each 0–20, plus team overall /80.
- **Player Stats:** original player boxes for projected players plus the uploaded projection and optimizer fields.
- **Player Rankings:** Paint/Shooting/Attacking/Fantasy each up to 30, Roto up to 50, and overall up to 170. Includes a breakdown of each category's raw-player score and opponent bonus, matchup coverage status, partial-data status, and downloadable leaderboard CSV.

## Validated in development

- Python syntax compilation passed.
- November 7 corrected workbook: 14 team rows and 202 player rows parsed.
- November 7 optimizer: all 2,400 roster selections matched to the projection-eligible player list, with two additional name aliases (Mikel Brown Jr. and Moritz Wagner) using team-aware matching.
- Projected-minutes thresholds, scoring maxima and unmatched-opponent partial-scoring behavior passed calculation tests.
- Tested projection header detection for `MINUTES`, `FPTS`, and `FPTS/$` as distinct columns.

These are local calculation tests, **not** a live Streamlit deployment test. Projection scores are custom ranking metrics, not guaranteed NBA results.
