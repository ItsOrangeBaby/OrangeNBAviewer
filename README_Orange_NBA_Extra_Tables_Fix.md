# Orange NBA Viewer — fix for extra tables in uploaded workbooks

## Problem

A workbook may contain the required `Team Stats` and `Player Stats` tables **plus additional calculation/reference sections on the same sheets**. Previous app versions scanned every nonblank first-column value after row 2. As a result, a supplemental table repeating the same teams produced:

`Unable to read this NBA workbook: Team Stats contains duplicate team names.`

The November 7, 2026 workbook contains the real 14-team table in **Team Stats rows 3–16**. Its supplemental **MATCHUP CALCULATION INPUTS** table starts in row 21 and repeats those teams. Its **Player Stats** sheet also includes a `SOURCE:` note after the player table.

## Fix

Replace the deployed `app.py` with the included script. It now stops at a following supplemental section heading, repeated header, or source note rather than treating those records as more teams/players. The workbook's data and appended calculation columns remain untouched. Actual duplicate teams within the primary source table are still reported as errors.

## Daily workflow (any NBA slate)

1. Deploy `app.py` to the `Orange-NBA-Viewer` GitHub repository and keep `requirements.txt`.
2. Reload/reboot your Streamlit app to run the new script.
3. Upload your NBA workbook (Excel `.xlsx` format), optimized lineups CSV, and Roto projections CSV. **Filenames are not checked.**
4. Select the correct NBA slate date and projection pool.
5. Only players on the selected projections list appear in player views. Missing player stats stay blank, while valid stats contribute to partial scores. Player matchup multipliers remain disabled.

## Verification with the uploaded November 7 files

- Excel workbook parsed: **14 teams** and **202 players**.
- Team Stats parsed: **24 statistics per team**, the standard four six-stat scoring boxes.
- All four Team Matchup Score categories computed for 14 teams.
- Optimized lineups CSV parsed: **300 rows**, eight roster positions.
- Roto projections CSV parsed: **182 rows**.
- Regression: previous November 6 workbook still parses as 18 teams and 260 players.
- Python syntax compilation passed. Live Streamlit browser deployment has not been tested.

## Troubleshooting

If you still see the **same** duplicate-team error after replacing the script, make sure the GitHub commit deployed successfully and the Streamlit app is using the correct `app.py`. You do not need to delete the supplemental sections from the Excel workbook.
