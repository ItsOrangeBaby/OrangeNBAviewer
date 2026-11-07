# Orange NBA Viewer — Daily CSV Upload Fix

This Streamlit app works for **any NBA slate date**. The date of the supplied Night files was used only for import testing. The player-score multiplier system remains **removed**.

## What's fixed

1. **Three upload boxes appear immediately, in any order.** The NBA `.xlsx` workbook is no longer required before you can select the two `.csv` files. Uploaded CSVs remain selected when you subsequently add the workbook.
2. **Separate upload destinations** are clearly labeled:
   - **NBA master workbook (.xlsx)** — your `Player Stats` and `Team Stats` two-sheet workbook.
   - **Optimized Lineups (.csv)** — a DraftKings export with `PG`, `SG`, `SF`, `PF`, `C`, `G`, `F`, `UTIL` columns, or other supported lineup layouts.
   - **Roto Projections (.csv)** — a player projection table with `PLAYER`, `TEAM`, `SALARY`, `FLOOR`, `CEIL`, `FPTS/$`, and `SMASH`.
3. **Night/LATE slate detection:** When the optimized lineup filename contains `Night` or `Late`, the projection dropdown automatically defaults to **LATE** when that pool exists. You can override this selection. The choice resets appropriately when new files are supplied; there is no date hardcoding.
4. **Name matching:** A new `Robert Williams` → `Robert Williams III` alias lets DraftKings lineup records match the corresponding projection player. Existing aliases remain.
5. **Missing stats:** Genuine blank NBA values stay blank. Available raw NBA statistics and any lineup points still count toward the player's score. If there are no NBA stats for a projected player, they can receive a Roto-only score.
6. **Player eligibility:** Only players in the selected projections pool appear in Player Stats and Player Rankings.

## Use it every slate day

1. Replace **`app.py`** in your GitHub `Orange-NBA-Viewer` repository with the supplied updated file. Keep **`requirements.txt`**.
2. Commit your change and open your deployed Streamlit app.
3. At the top of the page, choose your `.xlsx`, optimized lineups `.csv`, and projections `.csv` in **any order**. The file sizes are small; no renaming or file conversion is required.
4. Confirm the game date and selected NBA matchups. For a Night optimizer file, inspect the **Projection slate** selection: `LATE` will be preselected when available.
5. Review Player Stats and Player Rankings. Opponent-based multipliers remain disabled. Your defined Roto lineup bands remain unchanged.

**Important:** If you upload only the two CSVs, they will show as selected, but scoring requires the Excel workbook too. If the three file-upload boxes do not appear at the top, the deployed Streamlit app is still running the older `app.py` or a cached/previous deployment.

## Import test using supplied 10 PM Night files

- Optimized lineup export: **300** lineups, **2,400** player appearances.
- Roto projection CSV: **230** CSV rows in total, of which **180 MAIN**, **49 LATE**, and **one numeric footer row**. The footer is excluded from the eligible-player list.
- Selected Night/LATE pool: **49** eligible projected players (4 teams).
- Projection imports mapped: **49 / 49**.
- Lineup selections matched after new alias: **2,400 / 2,400**.
- Overall-score rows: **49 / 49**, including one player without an NBA workbook stats row who receives a Roto-only score.
- Source workbook tested: `NBA_Stats_Nov_06_2026_9_Game_Slate.xlsx` (18 teams, 260 players).

Python syntax and in-memory scoring/import tests passed. Live Community Cloud file-upload interaction and external schedule service were not tested here.
