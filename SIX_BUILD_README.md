# Orange NBA Viewer — Complete Six-Build Player Scoring (/312)

## Deployment

Replace the GitHub repository's `app.py` with the `app.py` in this ZIP. Keep `requirements.txt` (requires `streamlit` and `pandas`). Deploy normally and upload your workbook, optimized-lineups CSV, and player projections CSV using the same upload boxes. The README is documentation; Streamlit does not require you to replace the existing repository README.

## Score structure

Six player-ranking builds each award up to 52 points, with **48 scored inputs** overall and a **312-point maximum**. The first five categories—Interior, Corner, Attack, Chaos, and Star—are unchanged from the prior corrected scorecard. They each have three opposing-team stats at 0/1/2/4 (12 possible) and five player stats at 0/1/2/4/8 (40 possible). Attack excludes `3FGM %UAST`.

The sixth is now named **Fantasy Score** (previously displayed as an incomplete `Player Score`). It contains exactly these eight scoring inputs:

| Source / input | 0 points | 1 point | 2 points | 4 points | 8 points |
| --- | --- | --- | --- | --- | --- |
| Game Total | Below 216.5 | 216.5–225.4 | 225.5–234.4 | 234.5+ | — |
| Spread (absolute value) | 9.6+ | 5.6–9.5 | 2.6–5.5 | 2.5 or less | — |
| OPP Points + DREB + OREB + AST | 174.5 or less | 174.6–183.4 | 183.5–192.4 | 192.5+ | — |
| Projected Minutes - last-six MPG | Below 5.5 | 5.5–8.4 | 8.5–11.9 | 12.0–15.9 | 16.0+ |
| # of Lineups | 0 | 1–34 | 35–84 | 85–149 | 150+ |
| Player Salary | Above $9,900 | $7,501–9,900 | $5,401–7,500 | $3,601–5,400 | $3,600 or less |
| Price Per Dollar | 3.0 or less | 3.1–3.9 | 4.0–4.9 | 5.0–6.9 | 7.0+ |
| Projected Fantasy Points | Below 25.0 | 25.0–36.9 | 37.0–50.9 | 51.0–64.9 | 65.0+ |

All numeric values use Excel-style half-up rounding to one decimal **before scoring**, consistent with the five other builds. `# of Lineups` is the raw count from the optimizer CSV. Missing sources remain blank, and the leaderboard explicitly marks incomplete player scores.

### Data mapping

- `Game Total` and `Spread` come from the player's own team's ESPN odds fields in **Team Stats**. `Spread` is scored using its absolute value so favored and underdog teams with the same margin earn the same points. Expected team totals do **not** replace the actual `Game Total` field.
- `OPP PTS + OPP DREB + OPP OREB + OPP AST` comes from the **opposing** team, not the player's own team.
- Last-six-game `MPG` comes from the uploaded Player Stats workbook; projected `MINUTES`, `SALARY`, `FPTS/$`, and `FPTS` come from the chosen projections pool.
- Salary is inversely scored; the lower the salary, the higher the earned points.
- Smash, Floor, and Ceiling no longer earn points in the sixth category, although they can still appear as raw projections elsewhere.

### Boundary conventions

The user-provided grid had two small boundary ambiguities. This app resolves them transparently:

- **Spread 9.6:** scored 0 (the table says 9.5 and below earns at least 1).
- **Opponent aggregate 174.5:** scored 0 (the 0-point rule says 174.5 or less, even though the 1-point rule starts at 174.5).
- **Salary $9,901–$9,999:** scored 0, since the 1-point rule begins at $9,900 or lower.

All other bands follow the supplied thresholds exactly. Any one-decimal-rounding edge cases are handled by rounding first. Current NBA salaries typically use $100 increments.

## Preserved features

The daily Game Slate, Team Stats, Matchup /80, Alpha chart, seven Player Stats viewer sections (including `%PF` and ESPN expected totals), projection pool selection, and CSV name matching are unchanged. Only the sixth Player Rankings scoring definition, its detailed tab, and the overall leaderboard total were corrected.

## Tested

Python compilation passed. Local integration with the uploaded November 25 workbook and CSV files identified 18 teams, 261 workbook player records, 152 players in MAIN projections, and 2,400 matched optimizer appearances across 300 lineups. The calculation produced 48 input rows per player and correctly bounded every section at 52 / overall 312. The 64 explicit Fantasy tier boundary tests and synthetic maximum-tier test passed. All five earlier build scores matched the previous app exactly. The six-tab display logic was smoke-tested without a running Streamlit server. Live deployment remains untested.
