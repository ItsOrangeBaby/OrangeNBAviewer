# Orange NBA Viewer — Player Matchup Adjustments

## Deploy
1. Replace `app.py` in your GitHub repository with the supplied `app.py`.
2. Keep or add `requirements.txt` in the repository root.
3. Commit your changes. Streamlit Community Cloud will redeploy the app.
4. Upload an NBA master-format `.xlsx` workbook in the running app. It must contain the original **Player Stats** and **Team Stats** worksheets with their grouped headers.
5. Open **Player Matchup Adjustments**, next to **Team Stats**, **Matchup Score**, and **Player Stats**.

## Usage
- Choose **All four preset matchups** to match players automatically to the opposing defense for Brooklyn/Boston, Detroit/New York, Sacramento/San Antonio, and Portland/LA Lakers (where both sides are in the workbook).
- Alternatively, select **Choose an offensive team**, and choose the **opponent defense**. The opposing team's *opponent-allowed* columns determine that team's matchup tiers.
- **Player rankings** shows players ranked by the adjusted value of one statistic. **Player breakdown** shows original stat, combined multiplier, adjusted stat and each individual team factor for one player.
- All numerical player statistics remain unchanged in the uploaded workbook; the adjustments are calculated for display and CSV export only.

## Multiplier rule
For each player statistic, apply **every linked opposing-team stat exactly once**:
- Unfavorable: x0.5
- Neutral: x1.0
- Favorable: x1.5

**Adjusted value = player source value × (product of all applicable factors).** For instance, 4.0 RA FGA × 1.5 × 1.0 × 0.5 = 3.0 adjusted RA FGA.

The tier thresholds are identical to the previously supplied Team Stats cutoffs (favorable means the defense has allowed more of that stat). Do not interpret adjusted values as guaranteed or model-calibrated projections.

### Player stat / linked opponent defensive stats
| Player stat | Opposing-team fields |
|---|---|
| PPG | OPP PTS |
| 2PM | OPP 2PM; OPP AST |
| 2PA | OPP 2PM |
| RA FGM / RA FGA | OPP PTS PAINT; RA FGM; RA FGA |
| NRA FGM / NRA FGA | OPP PTS PAINT; NRA FGA |
| 3PM | OPP AST; ATB 3PA |
| 3PA | ATB 3PA |
| C3PM / C3PA | C3PA |
| MR FGM / MR FGA | MR FGA |
| FTA | OPP FTA |
| OREB | OPP OREB; OPP PTS 2ND CHANCE |
| DREB | OPP DREB; OPP PTS FB |
| AST | OPP AST |
| STL | OPP STL; OPP PTS OFF TOV |
| BLK | OPP BLK |

**Missing-data handling:** The original Player Stats workbook does **not** contain player Corner 3PA (`C3PA`), even though it contains opponent C3PA. Therefore player C3PA adjustment is shown but remains blank. Missing player baseline, missing opposing-team stats, and duplicate source conflicts never become numeric zero. Player `3PA` refers to total player three-point attempts, not ATB-specific attempts; your existing mapping applies the opponent ATB 3PA tier to that total.

No changes to `requirements.txt` or your two-tab workbook are necessary.
