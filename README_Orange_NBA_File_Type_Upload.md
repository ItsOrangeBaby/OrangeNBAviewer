# Orange NBA Viewer — file-content upload fix

This is a **reusable daily slate** version of Orange NBA Viewer. It does not use uploaded filenames to decide whether a file is allowed, what date is selected, or which projection pool should be used.

## Upload controls

Three uploaders are visible immediately, in any order:

1. **NBA stats workbook** — the actual Excel `.xlsx` data. Its file name can be anything. The workbook must still contain `Team Stats` (25 columns; four original scoring boxes) and `Player Stats` (56 columns; nine six-stat boxes). Box 5 on the Team Stats page is a placeholder until its source statistics are defined.
2. **Optimized lineups CSV** — comma-separated text with lineup slots such as PG, SG, SF, PF, C, G, F and UTIL (or the other existing supported layouts). Any filename is accepted.
3. **Roto projections CSV** — CSV text with player/team fields, salary, floor, ceiling, FPTS/$, and SMASH; MAIN/LATE pools supported. Any filename is accepted.

The file picker is unrestricted (`st.file_uploader(..., type=None)`). After upload, the app checks the **file contents** for valid Excel workbook or CSV data. This removes extension-based picker restrictions and does not rely on a specific file name.

If a file contains the wrong kind of data, the app displays an error next to its upload status. No file is matched by its name.

## Date and projection pool selection

- Select the **Slate date (U.S. Eastern)** in the Game slate section. It defaults to today's Eastern date and can be changed for any day. The app fetches matchups for this date, with manual matchup entry available if needed. Your workbook/CSV filenames do not change the date.
- The **Projection slate** option suggests MAIN or LATE based on **actual player names** found in the optimized lineup data and projection pools, not `Night`, `Late`, `Nov`, or a date in the filenames. You may override it.
- Only players in the selected projections pool are eligible for Player Stats and Player Rankings. Players without Excel NBA stats can still receive Roto scores if included in projections.

## Scoring preserved

- Team Stats highlighting and the four 24-stat Matchup Score categories.
- Original player category totals (Paint, Shooting, Attacking, Fantasy) and the 15-tier Roto lineup score.
- **No 0.5x or 1.5x player multipliers.**
- Missing NBA statistics remain blank, but other valid statistics still add to player totals; scoring status marks partial coverage.
- Player-name aliases, including Robert Williams / Robert Williams III.

## Deployment

1. Replace `app.py` in your `Orange-NBA-Viewer` GitHub repository with the file from this ZIP.
2. Keep or replace `requirements.txt` with the packaged version (Streamlit and pandas).
3. Commit to the branch used by Streamlit Community Cloud, and confirm the app's entrypoint is this `app.py`.
4. Reload or reboot the deployed application. You should see all three upload boxes under **Upload daily NBA files** immediately.
5. Upload each file in the correct box, regardless of its filename, then select the correct slate date and projection pool.

**If the Browse buttons still cannot select or transfer files:** The issue is not filename matching inside the Python app; it may be a stale deployment, browser/file-picker issue, or platform upload error. Confirm that `Upload daily NBA files` is shown, inspect Streamlit's deployed logs, and capture the exact error or a screenshot.

## Validation

Verified with the example files (and again with simulated entirely arbitrary names/other extensions):

- NBA workbook: 18 teams, 260 player rows.
- Optimized lineups: 300 rows / 2,400 roster selections.
- Projection pool: automatic content-derived **LATE**, 49 players.
- Matched lineup selections: 2,400 / 2,400.
- Player rankings computed: 49 / 49.
- Invalid Excel/CSV input is rejected based on bytes, not name.
- Python syntax and in-memory calculation checks passed.

The live Streamlit browser upload interaction has **not** been tested in this environment.
