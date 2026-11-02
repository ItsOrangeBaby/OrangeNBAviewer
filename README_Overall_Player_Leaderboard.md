# Orange NBA Viewer — Overall Player Matchup Leaderboard

## Install / update

1. Extract `Orange_NBA_Overall_Player_Leaderboard.zip` (or download `app.py` directly).
2. In the GitHub repository for your Streamlit app, replace the existing **`app.py`** with this version. Leave `requirements.txt` in the repository root.
3. Commit changes and allow Streamlit Community Cloud to redeploy.
4. Open the app, upload your NBA `.xlsx` file, and go to **Player Matchup Adjustments → Overall player leaderboard**.

The workbook is uploaded inside the running app, so it does not need to be stored in GitHub.

## What changed

The Player Matchup Adjustments section now contains **three** tabs:

- **Overall player leaderboard** (opens first): Rank, Player, Team, Opponent, Paint & 2PT subtotal, Shooting subtotal, Other subtotal, Overall Score, Stats Present, and Missing Stats. Sort by clicking a column heading, filter by team/name, and download the leaderboard CSV.
- **Individual stat rankings** (unchanged): Rank players by any one adjusted player statistic.
- **Player breakdown**: The full per-stat multiplier calculations, plus an individual player's combined score, rank, and completeness count.

Existing Team Stats, Matchup Score, Player Stats, team-tier colors, and Excel uploading are unchanged.

## Overall score formula

**Per-stat adjusted value = original player's source statistic × product of all its relevant opposing-team defensive multipliers.**

- Unfavorable defense matchup: **0.5×**
- Neutral defense matchup: **1.0×**
- Favorable defense matchup: **1.5×**

**Overall Score = sum of all 18 available adjusted player statistics.**

| Category subtotal | Stats summed |
|---|---|
| Paint & 2PT | 2PM, 2PA, RA FGM, RA FGA, NRA FGM, NRA FGA |
| Shooting | 3PM, 3PA, C3PM, MR FGM, MR FGA |
| Other | PPG, FTA, OREB, DREB, AST, STL, BLK |

**Why 18 rather than 19?** The master workbook's Player Stats sheet doesn't include player **C3PA**, so that one metric is excluded from *every* player's overall score for consistent comparison. C3PA remains visible as missing in the individual stat rankings and breakdown.

Missing values for any of the other **18** required stats cause the player's **Overall Score and Rank** to remain blank. Zero is a valid number and contributes zero. In the supplied eight-team example, **102 of 109 players** have complete overall scores; **7 remain unranked** because of additional missing values.

Ties share the same ranking number. Ranking is calculated across all players in the active selected matchups **before** search or team filters; filters do not reassign the ranks.

**Important:** This is a custom unweighted sum of volume stats (including overlapping measures such as points, made shots, and attempts), not a calibrated projection of NBA points or an official fantasy points formula. Different statistics are added as raw values and may affect the ranking disproportionately. If you later decide on a different scoring formula, it can be updated independently.

## Files

- `app.py` — Streamlit app (contains all previous features plus the overall leaderboard)
- `requirements.txt` — Python dependencies
- `README_Overall_Player_Leaderboard.md` — this guide
