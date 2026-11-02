# Orange NBA Viewer — Team Threshold Update

This version updates the six requested opponent defensive matchup thresholds. No changes to uploaded workbook layout are needed.

| Statistic | Unfavorable | Neutral | Favorable |
|---|---|---|---|
| OPP BLK | ≤ 4.0 | > 4.0 and < 6.0 | ≥ 6.0 |
| OPP DREB | < 29.0 | ≥ 29.0 and < 35.0 | ≥ 35.0 |
| ATB 3PM | ≤ 9.0 | > 9.0 and < 11.0 | ≥ 11.0 |
| OPP 2PM | < 25.0 | ≥ 25.0 and < 30.0 | ≥ 30.0 |
| OPP OREB | ≤ 11.0 | > 11.0 and < 14.0 | ≥ 14.0 |
| NRA FGM | ≤ 5.0 | > 5.0 and < 10.0 | ≥ 10.0 |

**Note:** Where you gave only one new boundary (e.g., OPP DREB favorable at 35), the other old boundary is kept. NRA FGM's previous unfavorable cutoff was <6; the revised cutoff of ≤5 makes values between 5.0 and 6.0 neutral.

All other team thresholds remain unchanged. The new boundaries apply consistently to: (a) Team Stats number colors, (b) Matchup Score 0/3/6 and 0/1/2 points, and (c) Player Matchup Adjustment multipliers and overall rankings. Uploaded Excel values are not changed. Only one decimal place is displayed.

## Deploy

1. Replace `app.py` in your GitHub `Orange-NBA-Viewer` repository with the one supplied here.
2. Keep `requirements.txt` at repository root.
3. Commit the change, let Streamlit redeploy, and refresh your app.
4. Upload your `.xlsx` file in the app.

Validated by Python syntax check, explicit boundary tests, and scoring against the supplied 30-team November 2 workbook. Live Streamlit interface not tested.
