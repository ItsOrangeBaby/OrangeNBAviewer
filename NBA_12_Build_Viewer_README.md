# Orange NBA Viewer — 12 Build /107.50

This is a date-independent, four-page Streamlit NBA viewer. Upload current-day data inside the app. File names do not determine the slate; the program reads the file columns and the selected projection pool.

## GitHub deployment

1. Replace your repository's `app.py` with the `app.py` in this ZIP.
2. Keep `requirements.txt` in the repository (or use the bundled file).
3. Push to GitHub and let Streamlit Community Cloud refresh.
4. Upload the daily NBA stats workbook, player projections CSV and optimized lineups CSV(s).
5. Choose MAIN or LATE projection pool, then choose the matching optimizer CSV. `Auto` selects the CSV with the strongest player overlap. You may upload one or both optimized-lineup files in a single uploader.
6. Filter to one game or the whole workbook slate.

## Pages

### 1. Team Essentials

- Scoring Essentials: Points, FGM, FTM, RA FGA, NRA FGA, MR FGA, ATB 3PTA, C3 PTA.
- Interior Essentials: Paint points, second-chance points, 2PTM, **2PTA derived as opponent FGA minus corner/above-the-break 3PA**, RA FGM, NRA FGM, FTA, offensive rebounds.
- Exterior Essentials: FGA, derived 2PTA, combined 3PTA, MR FGM, ATB 3PTM, corner 3PTM, AST + STL, DREB + BLK.
- Fantasy Outlook: sportsbook game total and team spread from the workbook; distinct optimizer players in the game; count of players projected for >=20 minutes; count projected for >=30 fantasy points; **sum of the players' projected CEIL values**. The last value is a summed projection, not a prediction of how many players hit their ceilings.

### 2. Matchup Score

Two charts of six team metrics each, as specified by the user. Each displays the actual opponent-allowed value, awarded scorecard points, a category total, and source coverage.

- Scoring /5.25: scorecard rule IDs **43, 7, 25, 19, 20, 77**.
- Production /6.25: rule IDs **74, 76, 55, 26, 13, 14**.
- Combined diagnostic maximum: **11.50**. These metrics also appear in other scorecard builds; this diagnostic is not an additional player-score bonus.

### 3. Player Stats

One section per scorecard build with raw/derived player metric values. Each of the first ten builds (Corner, Interior, Chaos, Attack, Star, Playmaker, Rebound, Upside, Fantasy, PRA) contains four player metrics. Roto contains ten player metrics; Matchup contains only team metrics.

### 4. Player Rankings

**12 builds / 107.50 maximum**, from the user's `NBA_12_Build_Blank_Scorecard_107_50(1).xlsx`. All 80 row thresholds (including quarter-point, 1.25, 1.5 and 2-point awards) are embedded in `app.py`. The scorecard itself does not need to be uploaded in the daily workflow.

- Ten 7.5-point builds: Corner, Interior, Chaos, Attack, Star, Playmaker, Rebound, Upside, Fantasy, PRA.
- Roto /20.
- Matchup /12.5.

The program uses the defending **opponent's** team statistics for each opposing-team metric, not the player's own defense. Game Total and absolute Spread use the player’s game, with spread taken from their workbook team row. Missing inputs remain blank and cause a Partial status; a real zero can score 0. Scores are assigned to the one-decimal-rounded input, using the exact original scorecard intervals. Calculations are not normalized upward to fill unavailable weights.

## Matching and stat calculations

- Team code variants like GS/GSW, NY/NYK, NO/NOP, PHO/PHX and SA/SAS are normalized. Named player aliases in the uploaded rules document are recognized, with no arbitrary fuzzy matches.
- The projections CSV determines the player roster. A projected player absent from the workbook remains visible with missing historical stats.
- Optimized lineup appearances are distinct per uploaded lineup row (eight DK slots); raw lineups stay separate from points.
- If both optimizer CSVs are uploaded, only one is used at a time to avoid counting main and night lineups together.
- Player 3PM = C3PM + ATB 3PM when no direct 3PM appears in the workbook.
- Opponent 3PTM = team C3PM + ATB 3PM; opponent 3PTA = C3PA + ATB 3PA. Opponent FTM is read directly from the extra Team Stats column.
- Minutes difference = CSV projected minutes minus last-six-game workbook MPG. Team OPP and player R-A values are not guessed from missing sources.
- All twelve requested Page 2 metrics have exact threshold matches in the uploaded scorecard.

## Local verification completed

- Python syntax compilation passed.
- Tested all 80 scorecard scoring rows with the actual scorecard tiers (no mismatches in the representative tier-boundary tests).
- Parsed the uploaded November 27 workbook: 22 teams and 322 player rows.
- MAIN: 181 projection players, all 2,400 optimized lineup appearances matched across 300 lineups; 172 players had complete 80-input data.
- LATE: 106 projection players, 2,399 of 2,400 appearances matched across 300 lineups; 102 players had complete data.
- Full interface-path smoke test passed with a Streamlit test stub. A live hosted Streamlit deployment has **not** been tested.
