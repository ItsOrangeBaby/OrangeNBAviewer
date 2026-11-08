# Orange NBA Viewer — Reusable Daily Slate (Roto /80, Overall /200)

This update is **not tied to any calendar date**. Keep `app.py` in your GitHub repository and upload each day's NBA workbook, optimized lineups CSV, and projections CSV to Streamlit.

## Installing this update

1. Replace the repository's existing `app.py` with the supplied version.
2. Keep `requirements.txt` in the repository (included in the ZIP).
3. Commit the change and wait for the Streamlit deployment to refresh.
4. Open the app and upload the day's files through the three upload boxes. Filenames are not used to reject uploads; the app validates content.
5. Select the game date and verify matchups or supply manual pairings if the schedule lookup cannot return them.
6. Select the relevant projection pool. Only players listed in the selected projection pool are eligible for the player views and rankings.

The original Player Stats, Team Stats, Team Matchup Score, automatic daily slate lookup, CSV name aliases, and additive *half of the opposing team's* matchup-category bonus remain unchanged. No player multipliers are used.

## Player scoring: 200 maximum

| Category | Base player metrics | Opposing team's matchup bonus | Category maximum |
|---|---:|---:|---:|
| Paint | 20 | 10 (half of Paint /20) | 30 |
| Shooting | 20 | 10 (half of Shooting /20) | 30 |
| Attacking | 20 | 10 (half of team's **Scoring** /20) | 30 |
| Fantasy | 20 | 10 (half of Fantasy /20) | 30 |
| Roto | — | — | **80** |
| **Overall** | | | **200** |

Each Roto component is scored independently. Blank/invalid source values remain blank and do not prevent the other components from contributing. A real reported zero earns zero points.

## Roto scoring: 80 maximum

| Component | Maximum | Projection/optimizer input |
|---|---:|---|
| Lineup appearances | 20 | optimizer CSV lineup counts |
| Smash | 10 | `SMASH` |
| Points per dollar | 10 | `FPTS/$` |
| Salary price | 10 | `SALARY` |
| Projected minutes | 10 | `MINUTES` |
| Projected floor | 5 | `FLOOR` |
| Projected ceiling | 5 | `CEIL` |
| Projected fantasy points | 10 | `FPTS` |
| **Total** | **80** | |

### Optimized lineup appearances

| Appearances | Points | Appearances | Points |
|---|---:|---|---:|
| 0 | 0 | 1–6 | 1 |
| 7–14 | 2 | 15–24 | 3 |
| 25–39 | 4 | 40–59 | 5 |
| 60–74 | 6 | 75–89 | 7 |
| 90–109 | 8 | 110–124 | 9 |
| 125–144 | 10 | 145–159 | 11 |
| 160–174 | 12 | 175–184 | 13 |
| 185–194 | 14 | 195–204 | 15 |
| 205–214 | 16 | 215–224 | 17 |
| 225–234 | 18 | 235–244 | 19 |
| 245+ | 20 | | |

### Points per dollar (`FPTS/$`)

| Minimum value | Points |
|---|---:|
| Below 2.6 | 0 |
| 2.6 | 1 |
| 3.3 | 2 |
| 3.8 | 3 |
| 4.3 | 4 |
| 4.7 | 5 |
| 5.2 | 6 |
| 5.5 | 7 |
| 6.4 | 8 |
| 6.9 | 9 |
| 7.0+ | 10 |

### Salary (`SALARY`)

| Salary band | Points |
|---|---:|
| $3,700 or below | 10 |
| $3,701–$4,200 | 9 |
| $4,201–$4,900 | 8 |
| $4,901–$5,600 | 7 |
| $5,601–$6,400 | 6 |
| $6,401–$7,200 | 5 |
| $7,201–$8,000 | 4 |
| $8,001–$9,000 | 3 |
| $9,001–$10,000 | 2 |
| $10,001–$10,500 | 1 |
| Above $10,500 | 0 |

The salary rule intentionally awards more points to cheaper players.

### Projected minutes (`MINUTES`)

| Projected minutes | Points |
|---|---:|
| Below 17 | 0 |
| 17 to below 20 | 1 |
| 20 to below 24 | 2 |
| 24 to below 27 | 3 |
| 27 to below 29 | 4 |
| 29 to below 31 | 5 |
| 31 to below 33 | 6 |
| 33 to below 35 | 7 |
| 35 to below 37 | 8 |
| 37 to below 40 | 9 |
| 40+ | 10 |

### Projected floor (`FLOOR`)

| Minimum floor | Points |
|---|---:|
| Below 24 | 0 |
| 24 | 1 |
| 29 | 2 |
| 34 | 3 |
| 39 | 4 |
| 40+ | 5 |

### Projected ceiling (`CEIL`)

| Minimum ceiling | Points |
|---|---:|
| Below 39 | 0 |
| 39 | 1 |
| 49 | 2 |
| 59 | 3 |
| 69 | 4 |
| 70+ | 5 |

### Projected fantasy points (`FPTS`)

| Minimum FPTS | Points |
|---|---:|
| Below 19 | 0 |
| 19 | 1 |
| 24 | 2 |
| 29 | 3 |
| 34 | 4 |
| 39 | 5 |
| 44 | 6 |
| 49 | 7 |
| 54 | 8 |
| 59 | 9 |
| 60+ | 10 |

### Smash (`SMASH`) — 10 points (Opto removed)

| Minimum Smash value | Points |
|---|---:|
| Below 9 | 0 |
| 9 | 1 |
| 14 | 2 |
| 19 | 3 |
| 24 | 4 |
| 29 | 5 |
| 34 | 6 |
| 39 | 7 |
| 44 | 8 |
| 49 | 9 |
| 50+ | 10 |

Each row is an inclusive minimum, so a Smash value of 13.9 earns 1 point, 49.9 earns 9 points, and 50 or greater earns 10 points. `OPTO` may exist in the uploaded projections CSV but is ignored entirely by Roto scoring and is not displayed.

**Explicit resolutions for remaining boundary gaps:**

- The instruction said `below 39 = 9` and `40+ = 10` projected minutes; 39 to below 40 is assigned 9, until the specified 40-point threshold is reached.
- The instruction said salary at most $10,500 earns 1 and $10,600+ earns 0; values above $10,500 and below $10,600 are assigned 0. (NBA DFS salaries commonly use $100 increments.)

These conventions can be revised on request; they are shown in the app's Roto threshold expander.

## Checks performed

- Script compiled successfully as Python.
- Tested all lineup tier boundaries (including 0, 245, and 300).
- Tested the new Smash 9/14/19/24/29/34/39/44/49/50 thresholds, and verified Opto is excluded.
- Tested every projection-stat minimum threshold, salary boundary, blank/invalid input handling, and maximum Roto = 80.
- Validated scoring against 181 player rows from the uploaded November 7 projection CSV; all calculated scores were in the 0–80 range. A max-tier synthetic player scores exactly 80, and missing fields produce a Partial score without discarding valid components.
- This is local calculation validation, **not** a test of the live Streamlit deployment.
