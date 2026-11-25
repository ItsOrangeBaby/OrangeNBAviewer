# Orange NBA Viewer — Final Six Player Builds

## Install

Replace the Streamlit repository's `app.py` with the file in this package. Keep your existing `requirements.txt` if it already contains `pandas` and `streamlit` (the included requirements file can also be used). Deploy the repository as usual. No extra upload box is required for the scoring card because the five completed builds' exact Numeric Thresholds are embedded in `app.py`.

Upload the three daily data files:

1. NBA workbook with the **Team Stats** and **Player Stats** sheets (Nov 25 workbook includes opponent FTM in the supplemental AL column).
2. Optimized lineups CSV (roster columns, including PG/SG/SF/PF/C/G/F/UTIL).
3. Player projections CSV (contains `PLAYER`, `TEAM`, `OPP`, `MINUTES`, `SMASH`, `FPTS/$`, `FLOOR`, `CEIL`, `FPTS`).

The existing Game Slate, Team Stats, Team Matchup /80, Alpha chart, and seven Player Stats views remain unchanged. **Player Rankings** has been replaced by one Overall Leaderboard and six build tabs.

## What is scored

These builds have the exact 3 opponent-team + 5 individual-player tier definitions from **NBA_Final_Six_Build_Scorecard_Attack_Corrected.xlsx**:

- Interior Score — maximum 52
- Corner Score — maximum 52
- Attack Score — maximum 52. The previously removed `3FGM %UAST` input is not scored.
- Chaos Score — maximum 52
- Star Score — maximum 52

Team tiers score **0 / 1 / 2 / 4**. Player tiers score **0 / 1 / 2 / 4 / 8**. Each metric is rounded to one decimal first using Excel-compatible half-up rounding. Combined statistics are computed only when all their source values are present. A missing field stays missing; zero is a real zero.

The sixth **Player Score** has one defined 4-point scoring row in the corrected scorecard:
`OPP PTS + OPP DREB + OPP OREB + OPP AST`, with 1/2/4 points beginning at 180/185/191 (all inclusive).

Its additional **seven CSV metrics** are available as **raw data**: number of optimizer lineups, Smash, FPTS/$, projected minutes minus last-six workbook MPG, Floor, Ceiling, and projected FPTS. Projected minutes, MPG, and the difference are all exposed. The corrected scorecard supplies **no scoring thresholds for these seven fields**, so the app does **not** invent any tiers, does **not** reuse old Roto /80 scores, and does **not** pretend that the sixth section reaches 52.

The defined scoring maximum is **264** (five 52-point builds + the sixth 4-point team metric). The desired completed 312-point ranking remains pending **48 points of unspecified sixth-section scoring rules**. All ranked totals visibly say `Defined Total / 264`.

## Source resolution

The updated November 25 Excel workbook includes `OPP FTM` directly in Team Stats' supplemental section. The program uses that source without estimating free throws. Player `3PM` is reconstructed as `C3PM + ATB 3PM` when necessary. Combined source terms are preserved and scored as their rounded sum.

The optimizer lineup CSV controls raw lineup appearances. The projection CSV controls eligible players. Game/opponent matching uses the user's selected games, falling back to the projection CSV's `OPP` when required.

## Tests performed

- Python syntax compilation.
- Threshold boundary tests for all 41 defined inputs.
- Imported November 25 workbook: 18 team rows, 261 player rows, including OPP FTM.
- Imported November 25 MAIN projections: 152 eligible players.
- Read the 300-lineup CSV: 2,400 eligible lineup appearances matched; six-build leaderboard generated for 152 players.
- Checked all category scores stay inside their defined maxima and missing inputs remain unavailable.
- Simulated six-tab display logic without a running Streamlit server.

Live Streamlit deployment was **not** exercised.
