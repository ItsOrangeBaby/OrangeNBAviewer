#!/usr/bin/env python3
"""Orange NBA Viewer — reusable daily NBA workbook and CSV uploads.

Deploy this file as app.py on Streamlit Community Cloud.
Run locally: streamlit run app.py

Does not require any workbook to be stored in GitHub. The file you upload stays
in the current Streamlit session; it is parsed in memory and is not saved.
"""

from __future__ import annotations

import pandas as pd
import io
import math
import re
import zipfile
import json
import hashlib
from datetime import date, datetime
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo
from xml.etree import ElementTree as ET

import streamlit as st

TEAM_CODES = {
    "Atlanta Hawks": "ATL", "Boston Celtics": "BOS", "Brooklyn Nets": "BKN",
    "Charlotte Hornets": "CHA", "Chicago Bulls": "CHI", "Cleveland Cavaliers": "CLE",
    "Dallas Mavericks": "DAL", "Denver Nuggets": "DEN", "Detroit Pistons": "DET",
    "Golden State Warriors": "GSW", "Houston Rockets": "HOU", "Indiana Pacers": "IND",
    "LA Clippers": "LAC", "Los Angeles Clippers": "LAC", "Los Angeles Lakers": "LAL",
    "Memphis Grizzlies": "MEM", "Miami Heat": "MIA", "Milwaukee Bucks": "MIL",
    "Minnesota Timberwolves": "MIN", "New Orleans Pelicans": "NOP",
    "New York Knicks": "NYK", "Oklahoma City Thunder": "OKC", "Orlando Magic": "ORL",
    "Philadelphia 76ers": "PHI", "Phoenix Suns": "PHX",
    "Portland Trail Blazers": "POR", "Sacramento Kings": "SAC",
    "San Antonio Spurs": "SAS", "Toronto Raptors": "TOR", "Utah Jazz": "UTA",
    "Washington Wizards": "WAS",
}

# Thresholds apply to opponent statistics ALLOWED by the defending team.
# Higher allowed = a more FAVORABLE matchup for the opposing offense.
# Use the beginning of each stated integer range as an inclusive cutoff:
# e.g., <=24 / 25-30 / >=31 means value <25 / 25<=value<31 / value>=31.
# This also covers decimals such as 24.7 and 30.8 without gaps.
# Tuple: (neutral_start, favorable_start, displayed_average_if_provided)
TEAM_MATCHUP_THRESHOLDS = {
    "OPP 2PM": (25, 30, None),
    "OPP PTS PAINT": (43, 54, None),
    "OPP PTS 2ND CHANCE": (14, 19, None),
    "OPP OREB": (11, 14, None),
    "RA FGM": (14, 20, None),
    "OPP FTA": (22, 29, None),

    "OPP FGM": (39, 44, 40.94),
    "OPP AST": (24, 29, 26.09),
    "NRA FGM": (6, 10, 7.78),
    "MR FGM": (3, 5, 3.56),
    "C3PM": (3, 5, 3.46),
    "ATB 3PM": (9, 11, 9.47),

    "OPP FGA": (86, 93, 89.13),
    "OPP PTS OFF TOV": (17, 22, 19.09),
    "OPP PTS FB": (13, 19, 15.18),
    "OPP DREB": (29, 35, 32.40),
    "OPP STL": (7, 11, 8.78),
    "OPP BLK": (4, 6, 5.17),
    # Overall opponent 3PM is derived from ATB and corner makes if not supplied.
    "OPP 3PM": (13, 15, None),

    "OPP PTS": (106, 124, 114.68),
    "RA FGA": (21, 30, 25.15),
    "NRA FGA": (14, 20, 17.46),
    "MR FGA": (6, 11, 8.89),
    "C3PA": (9, 12, 10.08),
    "ATB 3PA": (24, 31, 27.54),
}

# Exact user-defined "X or lower" cutoffs (inclusive). Unlike the old integer
# tiers, decimals immediately above X are neutral, not unfavorable. All other
# statistics preserve their original start-of-neutral inclusive cutoff.
TEAM_UNFAVORABLE_MAX_INCLUSIVE = {
    "OPP BLK": 4,
    "OPP OREB": 11,
    "ATB 3PM": 9,
    "NRA FGM": 5,
    "OPP 3PM": 12,
}


def tier_ranges(field: str) -> tuple[str, str, str]:
    """Human-readable tier boundaries, consistent with decimal scoring."""
    neutral_start, favorable_start, _ = TEAM_MATCHUP_THRESHOLDS[field]
    if field in TEAM_UNFAVORABLE_MAX_INCLUSIVE:
        maximum = TEAM_UNFAVORABLE_MAX_INCLUSIVE[field]
        return (f"≤ {maximum}", f"> {maximum} to < {favorable_start}", f"≥ {favorable_start}")
    return (f"< {neutral_start}", f"{neutral_start} to < {favorable_start}", f"≥ {favorable_start}")


# Color ONLY the numeric text. Do not set a background color: the table
# retains the Streamlit theme's normal background in both dark and light mode.
TIER_STYLES = {
    "Unfavorable": "color: #E45C66; font-weight: 700;",
    "Neutral": "color: #DBA829; font-weight: 700;",
    "Favorable": "color: #25AD68; font-weight: 700;",
}


def matchup_tier(field: str, value: object) -> str | None:
    """Classify a defensive stat allowed; missing values stay unclassified."""
    thresholds = TEAM_MATCHUP_THRESHOLDS.get(field)
    if thresholds is None or pd.isna(value):
        return None
    try:
        number = float(value)
    except (ValueError, TypeError):
        return None
    if not math.isfinite(number):
        return None
    neutral_start, favorable_start, _average = thresholds
    if number >= favorable_start:
        return "Favorable"
    if field in TEAM_UNFAVORABLE_MAX_INCLUSIVE:
        return ("Unfavorable" if number <= TEAM_UNFAVORABLE_MAX_INCLUSIVE[field]
                else "Neutral")
    if number >= neutral_start:
        return "Neutral"
    return "Unfavorable"


def style_team_cell(value: object, field: str) -> str:
    tier = matchup_tier(field, value)
    return TIER_STYLES.get(tier, "")


def threshold_reference(fields: list[str]) -> pd.DataFrame:
    """Display the supplied tiers without altering the uploaded statistics."""
    rows = []
    for field in fields:
        _low_end, _high_start, average = TEAM_MATCHUP_THRESHOLDS[field]
        unfavorable, neutral, favorable = tier_ranges(field)
        row = {
            "Statistic": field,
            "Unfavorable": unfavorable,
            "Neutral": neutral,
            "Favorable": favorable,
        }
        if average is not None:
            row["Provided NBA avg"] = round(float(average), 1)
        rows.append(row)
    return pd.DataFrame(rows)


def one_decimal_config(frame: pd.DataFrame) -> dict:
    """Streamlit column config that shows numeric columns with one decimal place."""
    config = {}
    for column in frame.columns:
        if pd.api.types.is_numeric_dtype(frame[column]):
            config[column] = st.column_config.NumberColumn(column, format="%.1f")
    return config


# MATCHUP SCORE — FOUR CATEGORIES OF FIVE STATS, EACH WORTH 20 POINTS.
# A stat earns exactly 0, 1, 2 or 4 points depending on the ACTUAL defense-
# allowed value. Thresholds use inclusive minimums; tier values persist up to
# the next minimum (even if a description lists a smaller decimal upper bound).
# Keep these four-tier scores separate from the original three-tier coloring
# and Alpha /50 model, which retain their established legacy thresholds.
# Each entry: display label, source workbook header, maximum points.
MATCHUP_SCORE_GROUPS = {
    "Interior Score": [
        ("2PTM", "OPP 2PM", 4),
        ("Points In Paint", "OPP PTS PAINT", 4),
        ("RA FGA", "RA FGA", 4),
        ("NRA FGA", "NRA FGA", 4),
        ("FTA", "OPP FTA", 4),
    ],
    "Exterior Score": [
        ("FGA", "OPP FGA", 4),
        ("3PTM", "OPP 3PM", 4),
        ("MR FGA", "MR FGA", 4),
        ("C3PA", "C3PA", 4),
        ("ATB 3PA", "ATB 3PA", 4),
    ],
    "Production Score": [
        ("FGM", "OPP FGM", 4),
        ("Defensive Rebound", "OPP DREB", 4),
        ("Assists", "OPP AST", 4),
        ("Steals", "OPP STL", 4),
        ("Blocks", "OPP BLK", 4),
    ],
    "Chaos Score": [
        ("Points", "OPP PTS", 4),
        ("Offensive Rebounds", "OPP OREB", 4),
        ("Second Chance Points", "OPP PTS 2ND CHANCE", 4),
        ("Fastbreak Points", "OPP PTS FB", 4),
        ("Points Off Turnovers", "OPP PTS OFF TOV", 4),
    ],
}

# Inclusive starts of the 1-point, 2-point and 4-point tiers, respectively.
# All source values below the first cutoff receive zero. Values with more
# precision than one decimal are scored by these exact numeric thresholds.
MATCHUP_SCORE_CUTOFFS = {
    "OPP 2PM": (25.1, 27.0, 30.0),
    "OPP PTS PAINT": (43.1, 47.5, 52.0),
    "RA FGA": (20.1, 24.5, 30.0),
    "NRA FGA": (12.6, 17.0, 20.0),
    "OPP FTA": (19.6, 23.5, 27.0),

    "OPP FGA": (84.6, 88.5, 90.5),
    "OPP 3PM": (12.1, 14.0, 15.5),
    "MR FGA": (7.1, 8.5, 11.0),
    "C3PA": (8.6, 10.0, 12.5),
    "ATB 3PA": (25.1, 28.0, 29.5),

    "OPP FGM": (38.1, 41.0, 43.0),
    "OPP DREB": (28.6, 31.5, 34.5),
    "OPP AST": (24.1, 26.5, 29.0),
    "OPP STL": (6.6, 8.0, 9.5),
    "OPP BLK": (4.1, 5.0, 6.0),

    "OPP PTS": (108.1, 113.6, 120.0),
    "OPP OREB": (10.1, 12.0, 14.0),
    "OPP PTS 2ND CHANCE": (12.1, 15.5, 17.5),
    "OPP PTS FB": (12.1, 15.5, 17.5),
    "OPP PTS OFF TOV": (15.1, 18.0, 21.0),
}
assert len(MATCHUP_SCORE_GROUPS) == 4
assert all(len(group) == 5 and sum(weight for _, _, weight in group) == 20
           for group in MATCHUP_SCORE_GROUPS.values())
assert set(MATCHUP_SCORE_CUTOFFS) == {
    field for group in MATCHUP_SCORE_GROUPS.values() for _, field, _ in group
}
assert len(MATCHUP_SCORE_CUTOFFS) == 20


def matchup_score_points(field: str, value: object) -> int | None:
    """New /4 scoring: four exact tiers, no rounding or fabricated zeros."""
    number = valid_number_or_nan(value)
    if pd.isna(number):
        return None
    one_start, two_start, four_start = MATCHUP_SCORE_CUTOFFS[field]
    if number >= four_start:
        return 4
    if number >= two_start:
        return 2
    if number >= one_start:
        return 1
    return 0


def matchup_score_tier_style(field: str, value: object) -> str:
    """Color raw defense-allowed values to match their new four-tier points."""
    score = matchup_score_points(field, value)
    return MATCHUP_POINT_STYLES.get(score, "")


MATCHUP_POINT_STYLES = {
    0: "color: #E45C66; font-weight: 700;",  # unfavorable
    1: "color: #D9813B; font-weight: 700;",  # slight opportunity
    2: "color: #DBA829; font-weight: 700;",  # positive opportunity
    4: "color: #25AD68; font-weight: 700;",  # favorable
}


# Alpha is 12 actual defense-allowed stats, with a derived rating /50.
# No Alpha weights were supplied: its first two stats are weighted /5,
# and the remaining ten /4 (2*5 + 10*4 = 50). These can be tuned later.
ALPHA_STATS = [
    ("Points", "OPP PTS", 5),
    ("OPP FGM", "OPP FGM", 5),
    ("RA FGA", "RA FGA", 4),
    ("NRA FGA", "NRA FGA", 4),
    ("FTA", "OPP FTA", 4),
    ("MR FGA", "MR FGA", 4),
    ("C3PTA", "C3PA", 4),
    ("ATB 3PTA", "ATB 3PA", 4),
    ("OREB", "OPP OREB", 4),
    ("AST", "OPP AST", 4),
    ("Steals", "OPP STL", 4),
    ("Blocks", "OPP BLK", 4),
]
assert sum(maximum for _, _, maximum in ALPHA_STATS) == 50

# The zero box is a quick-view table, not an additional 20-point category.
TEAM_BOX_ZERO_STATS = [
    ("Points", "OPP PTS"),
    ("OPP FGA", "OPP FGA"),
    ("RA FGM", "RA FGM"),
    ("NRA FGM", "NRA FGM"),
    ("MR FGM", "MR FGM"),
    ("C3PM", "C3PM"),
    ("ATB 3PM", "ATB 3PM"),
    ("OPP FTA", "OPP FTA"),
    ("OREB", "OPP OREB"),
    ("Steals", "OPP STL"),
]


def combined_team_3pm(values: dict) -> float:
    """Use actual opponent makes when present, otherwise both location sums.

    A missing side of the split must not be silently replaced with zero.
    """
    direct = valid_number_or_nan(values.get("OPP 3PM"))
    if pd.notna(direct):
        return direct
    corner = valid_number_or_nan(values.get("C3PM"))
    above = valid_number_or_nan(values.get("ATB 3PM"))
    return corner + above if pd.notna(corner) and pd.notna(above) else float("nan")


def team_stat_value(values: dict, field: str) -> float:
    return (combined_team_3pm(values) if field == "OPP 3PM"
            else valid_number_or_nan(values.get(field)))


def team_values_for_scoring(sheet_data: dict) -> dict[str, dict]:
    """Flatten team boxes, including explicitly supplied supplemental stats."""
    all_fields = [field for box in sheet_data["boxes"] for field in box["fields"]]
    required_fields = [field for specs in MATCHUP_SCORE_GROUPS.values()
                       for _label, field, _maximum in specs if field != "OPP 3PM"]
    required_fields.extend(field for _, field, _ in ALPHA_STATS)
    required_fields.extend(field for _, field in TEAM_BOX_ZERO_STATS)
    missing = sorted(set(required_fields) - set(all_fields))
    if missing:
        raise ValueError("Team Stats is missing score fields: " + ", ".join(missing))
    if len(all_fields) != len(set(all_fields)):
        raise ValueError("Team Stats includes duplicate statistic headers.")
    return {
        entry["name"]: {
            **dict(zip(all_fields,
                       [value for group in entry["groups"] for value in group])),
            **entry.get("_team_projection", {}),
        }
        for entry in sheet_data["entries"]
    }


def matchup_points(field: str, value: object, maximum: int) -> float | None:
    """Classify 0 / half / full for /2, /4, /6 and Alpha /5 scores."""
    tier = matchup_tier(field, value)
    if tier is None:
        return None
    if maximum not in (2, 4, 5, 6):
        raise ValueError(f"Unsupported score maximum: {maximum}")
    return {"Unfavorable": 0.0, "Neutral": maximum / 2.0,
            "Favorable": float(maximum)}[tier]


def build_matchup_score_table(all_teams: dict[str, dict], group_name: str) -> pd.DataFrame:
    """Five distinct 0/1/2/4 metrics per category; report incomplete inputs."""
    specs = MATCHUP_SCORE_GROUPS[group_name]
    if len(specs) != 5 or len({field for _, field, _ in specs}) != 5:
        raise ValueError(f"{group_name} needs five unique team stats.")
    rows = []
    for team, values in all_teams.items():
        row = {"Team": team}
        for label, source_field, _maximum in specs:
            row[label] = matchup_score_points(source_field, team_stat_value(values, source_field))
        available = [row[label] for label, _, _ in specs if row[label] is not None]
        missing = [label for label, _, _ in specs if row[label] is None]
        row["Total / 20"] = sum(available) if available else None
        row["Stats Present"] = len(available)
        row["Status"] = "Complete" if not missing else ("Partial" if available else "Unavailable")
        row["Missing"] = ", ".join(missing)
        rows.append(row)
    result = pd.DataFrame(rows)
    for col in [label for label, _, _ in specs] + ["Total / 20", "Stats Present"]:
        result[col] = pd.to_numeric(result[col], errors="coerce")
    return result


def build_overall_score_table(all_teams: dict[str, dict]) -> pd.DataFrame:
    """Aggregate 20 unique defense-allowed metrics into four /20 scores."""
    specs = [field for group in MATCHUP_SCORE_GROUPS.values() for _, field, _ in group]
    if len(specs) != 20 or len(set(specs)) != 20:
        raise ValueError("Team Matchup Score must use 20 unique scoring stats.")
    by_group = {group: build_matchup_score_table(all_teams, group).set_index("Team")
                for group in MATCHUP_SCORE_GROUPS}
    rows = []
    for team in all_teams:
        row = {"Team": team}
        totals = []
        present = 0
        missing = []
        for group, table in by_group.items():
            score = table.at[team, "Total / 20"]
            row[group] = score
            if pd.notna(score):
                totals.append(float(score))
            present += int(table.at[team, "Stats Present"])
            if table.at[team, "Missing"]:
                missing.append(f"{group}: {table.at[team, 'Missing']}")
        row["Overall / 80"] = sum(totals) if totals else None
        row["Stats Present"] = present
        row["Status"] = "Complete" if present == 20 else ("Partial" if present else "Unavailable")
        row["Missing"] = "; ".join(missing)
        rows.append(row)
    return pd.DataFrame(rows)


def point_text_style(value: object, maximum: int) -> str:
    """Map the new 0 / 1 / 2 / 4 score to its exact tier color."""
    if pd.isna(value):
        return ""
    return MATCHUP_POINT_STYLES.get(int(value), "")


def matchup_score_threshold_reference(group_name: str) -> pd.DataFrame:
    """Show exact four-tier cutoffs for each category's five components."""
    rows = []
    for label, source_field, maximum in MATCHUP_SCORE_GROUPS[group_name]:
        one, two, four = MATCHUP_SCORE_CUTOFFS[source_field]
        rows.append({
            "Statistic": label,
            "Workbook source": source_field,
            "0 points": f"< {one:g}",
            "1 point": f"{one:g} to < {two:g}",
            "2 points": f"{two:g} to < {four:g}",
            "4 points": f"≥ {four:g}",
        })
    return pd.DataFrame(rows)


def show_score_table(frame: pd.DataFrame, group_name: str) -> None:
    """One five-stat score table and its four-tier threshold reference."""
    if frame.empty:
        st.info("No matching teams. Select another matchup or team filter.")
        return
    specs = MATCHUP_SCORE_GROUPS[group_name]
    stat_fields = [label for label, _, _ in specs]
    left, right = st.columns([2, 1])
    with left:
        sort_column = st.selectbox("Sort by", ["Total / 20", *stat_fields, "Stats Present", "Team"],
                                   key=f"score_{group_name}_sort")
    with right:
        sort_direction = st.selectbox("Order", ["Highest first", "Lowest first"],
                                      key=f"score_{group_name}_direction")
    result = frame.sort_values(sort_column,
                               ascending=sort_direction == "Lowest first",
                               kind="stable", na_position="last").reset_index(drop=True)
    styled = result.style
    for label, _field, maximum in specs:
        styled = styled.map(lambda value, m=maximum: point_text_style(value, m), subset=[label])
    styled = styled.map(lambda value: "font-weight: 700;" if pd.notna(value) else "",
                        subset=["Total / 20"])
    styled = styled.format({col: "{:.0f}" for col in stat_fields + ["Total / 20", "Stats Present"]},
                           na_rep="—")
    st.caption(f"{len(result)} teams · 5 scoring stats · maximum 20. Four tiers: "
               "0 / 1 / 2 / 4 points. Partial means at least one input is missing.")
    st.dataframe(styled, hide_index=True, use_container_width=True,
                 height=min(640, 80 + 35 * len(result)))
    with st.expander("See point thresholds and source columns"):
        st.dataframe(matchup_score_threshold_reference(group_name),
                     hide_index=True, use_container_width=True)
        st.caption("Every cutoff is inclusive. A value between listed thresholds "
                   "earns the highest tier it has reached; decimals are never rounded first.")
    st.download_button("Download score table CSV", result.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"orange_nba_{group_name.lower().replace(' ', '_')}.csv",
                       mime="text/csv", key=f"score_{group_name}_csv")


def build_alpha_table(teams: dict[str, dict], selected: list[str]) -> pd.DataFrame:
    """Display all 12 raw Alpha stats, then their 50-point tier-weighted sum."""
    rows = []
    for team in selected:
        if team not in teams:
            continue
        row = {"Team": team}
        earned = []
        for label, field, maximum in ALPHA_STATS:
            raw = team_stat_value(teams[team], field)
            row[label] = raw
            points = matchup_points(field, raw, maximum)
            if points is not None:
                earned.append(points)
        row["Total / 50"] = sum(earned) if earned else float("nan")
        rows.append(row)
    return pd.DataFrame(rows, columns=["Team", *(label for label, _, _ in ALPHA_STATS), "Total / 50"])


def show_alpha_table(teams: dict[str, dict], selected: list[str]) -> None:
    alpha = build_alpha_table(teams, selected)
    if alpha.empty:
        st.info("No teams selected for the Alpha chart.")
        return
    alpha = alpha.sort_values("Total / 50", ascending=False, na_position="last")
    styled = alpha.style
    for label, field, _maximum in ALPHA_STATS:
        styled = styled.map(lambda v, f=field: style_team_cell(v, f), subset=[label])
    styled = styled.format({label: "{:.1f}" for label, _, _ in ALPHA_STATS} |
                           {"Total / 50": "{:.1f}"}, na_rep="—")
    st.dataframe(styled, hide_index=True, use_container_width=True,
                 height=min(680, 85 + 35 * len(alpha)))
    st.caption("Alpha cells show actual defensive statistics, not score points. "
               "Green = favorable for the opponent, yellow = neutral, red = unfavorable. "
               "Total /50 uses 5-point tiers for Points and OPP FGM, and 4-point "
               "tiers for the other 10 stats. Missing values earn no assumed zero.")
    st.download_button("Download Alpha chart CSV", alpha.to_csv(index=False).encode("utf-8-sig"),
                       file_name="orange_nba_alpha.csv", mime="text/csv")


def show_team_box_zero(teams: dict[str, dict], selected: list[str]) -> None:
    rows = [{"Team": name, **{label: team_stat_value(teams[name], field)
                               for label, field in TEAM_BOX_ZERO_STATS}}
            for name in selected if name in teams]
    if not rows:
        st.info("No teams selected.")
        return
    frame = pd.DataFrame(rows)
    col_a, col_b = st.columns([2, 1])
    with col_a:
        sort_key = st.selectbox("Sort by", list(frame.columns[1:]) + ["Team"],
                                key="team_box_zero_sort")
    with col_b:
        sort_order = st.selectbox("Order", ["Highest first", "Lowest first"],
                                  key="team_box_zero_order")
    frame = frame.sort_values(sort_key, ascending=sort_order == "Lowest first",
                              na_position="last", kind="stable")
    styled = frame.style
    for label, field in TEAM_BOX_ZERO_STATS:
        styled = styled.map(lambda v, f=field: style_team_cell(v, f), subset=[label])
    st.dataframe(styled.format({label: "{:.1f}" for label, _ in TEAM_BOX_ZERO_STATS},
                               na_rep="—"), hide_index=True, use_container_width=True,
                 height=min(680, 85 + 35 * len(frame)))
    st.download_button("Download Team Box Zero CSV",
                       frame.to_csv(index=False).encode("utf-8-sig"),
                       file_name="orange_nba_team_box_zero.csv", mime="text/csv")


# PLAYER SCORING — ORIGINAL STATISTICS → 20-POINT CATEGORY SCORES
# No opponent-based multipliers. Original player statistics remain unchanged.
# For each stat, compare eligible players on the selected slate using percentile
# thirds: low / middle / high. The first two category stats are priority (0/3/6),
# the remaining four are support (0/1/2), for a 20-point category maximum.
# Availability is detected from each uploaded workbook. Modern workbooks can
# supply player C3PA; older workbooks may not. Do not hard-exclude valid data.
# The full list also preserves some additional, non-scoring stats in Player Detail.
PLAYER_DETAIL_STATS = (
    "PPG", "FTM", "2PM", "2PA", "RA FGM", "RA FGA", "NRA FGM", "NRA FGA",
    "3PM", "3PA", "C3PM", "C3PA", "ATB 3PM", "MR FGM", "MR FGA",
    "FTA", "OREB", "DREB", "AST", "STL", "BLK", "FGM", "FGA",
    "%OREB", "%DREB", "%AST", "%PF", "PF", "USG%", "FGM %UAST", "DD2", "TD3",
)
PLAYER_SCORE_CATEGORIES = {
    "Paint": ("2PM", "2PA", "RA FGM", "RA FGA", "NRA FGM", "NRA FGA"),
    "Shooting": ("3PM", "3PA", "C3PM", "C3PA", "MR FGM", "MR FGA"),
    "Attacking": ("USG%", "PPG", "FGM", "FGA", "%OREB", "FGM %UAST"),
    "Fantasy": ("STL", "BLK", "%DREB", "%AST", "DD2", "TD3"),
}
# As with Team Matchup Score: first two priority stats have a 6-point ceiling,
# and the next four have a 2-point ceiling. Percentile thresholds are derived
# from the eligible players, not fabricated absolute NBA player thresholds.
PLAYER_SCORE_WEIGHTS = {
    category: {field: (6 if i < 2 else 2) for i, field in enumerate(fields)}
    for category, fields in PLAYER_SCORE_CATEGORIES.items()
}
assert all(sum(weights.values()) == 20 for weights in PLAYER_SCORE_WEIGHTS.values())

# Player base-score category names are unchanged. Each receives HALF of
# the corresponding new 20-point score from the OPPOSING defense.
# Paint -> Interior, Shooting -> Exterior, Attacking -> Chaos,
# Fantasy -> Production (rebounding, assists, steals and blocks).
PLAYER_TO_TEAM_MATCHUP = {
    "Paint": "Interior Score",
    "Shooting": "Exterior Score",
    "Attacking": "Chaos Score",
    "Fantasy": "Production Score",
}
PLAYER_BASE_CATEGORY_MAX = 20
PLAYER_MATCHUP_BONUS_MAX = 10  # 50% of one team matchup category (0–20)
PLAYER_CATEGORY_MAX = PLAYER_BASE_CATEGORY_MAX + PLAYER_MATCHUP_BONUS_MAX


# Include all 24 stats in scoring input. If a whole stat is absent from an
# uploaded workbook, redistribute only that universally unavailable stat's
# capacity across the others in its category. Never estimate missing values.
OVERALL_SCORE_STATS = tuple(dict.fromkeys(
    field for group in PLAYER_SCORE_CATEGORIES.values()
    for field in group
))

# A genuine zero-lineup count only occurs when an optimizer CSV was uploaded.
# No filename restrictions; the lineup CSV defines actual appearance counts.
ROTO_LINEUP_TIERS = [
    (1, 6, 1), (7, 14, 2), (15, 24, 3), (25, 39, 4),
    (40, 59, 5), (60, 74, 6), (75, 89, 7), (90, 109, 8),
    (110, 124, 9), (125, 144, 10), (145, 159, 11),
    (160, 174, 12), (175, 184, 13), (185, 194, 14),
    (195, 204, 15), (205, 214, 16), (215, 224, 17),
    (225, 234, 18), (235, 244, 19), (245, float("inf"), 20),
]


def roto_points(appearances: object) -> float:
    count = valid_number_or_nan(appearances)
    if pd.isna(count) or count < 0 or int(count) != count:
        return float("nan")
    if count == 0:
        return 0.0
    return float(next(points for low, high, points in ROTO_LINEUP_TIERS if low <= count <= high))


# Inclusive minimum cutoffs: each value receives its highest attained tier.
# Smash now uses a 10-point ladder, with 9 (inclusive) as the first tier.
# Minutes 39.0–39.999 scores 9; the 10-point tier begins at 40.
# CSV percentages are supplied on a 0–100 scale, not 0–1.
ROTO_PROJECTION_TIERS = {
    "Smash": [
        (9, 1), (14, 2), (19, 3), (24, 4), (29, 5),
        (34, 6), (39, 7), (44, 8), (49, 9), (50, 10),
    ],
    "Price per Dollar": [
        (2.6, 1), (3.3, 2), (3.8, 3), (4.3, 4), (4.7, 5),
        (5.2, 6), (5.5, 7), (6.4, 8), (6.9, 9), (7, 10),
    ],
    "Projected Minutes": [
        (17, 1), (20, 2), (24, 3), (27, 4), (29, 5),
        (31, 6), (33, 7), (35, 8), (37, 9), (40, 10),
    ],
    "Projected Floor": [(24, 1), (29, 2), (34, 3), (39, 4), (40, 5)],
    "Projected Ceiling": [(39, 1), (49, 2), (59, 3), (69, 4), (70, 5)],
    "Projected Fantasy Points": [
        (19, 1), (24, 2), (29, 3), (34, 4), (39, 5),
        (44, 6), (49, 7), (54, 8), (59, 9), (60, 10),
    ],
}

# Salary rewards lower-priced players: first qualifying upper bound wins.
# $10,501–$10,599 is scored 0 because the last awarded bound is $10,500.
ROTO_SALARY_TIERS = [
    (3700, 10), (4200, 9), (4900, 8), (5600, 7), (6400, 6),
    (7200, 5), (8000, 4), (9000, 3), (10000, 2), (10500, 1),
]
ROTO_COMPONENT_LABELS = [
    "Lineup Points", "Smash Points", "PPD Points",
    "Salary Points", "Projected Minutes Points", "Floor Points",
    "Ceiling Points", "Projected FPTS Points",
]
ROTO_COMPONENT_SOURCES = {
    "Smash Points": "Smash",
    "PPD Points": "Price per Dollar",
    "Salary Points": "Player Price",
    "Projected Minutes Points": "Projected Minutes",
    "Floor Points": "Projected Floor",
    "Ceiling Points": "Projected Ceiling",
    "Projected FPTS Points": "Projected Fantasy Points",
}
ROTO_MAX = 20 + 10 + 10 + 10 + 10 + 5 + 5 + 10
PLAYER_OVERALL_MAX = 4 * PLAYER_CATEGORY_MAX + ROTO_MAX
assert ROTO_MAX == 80 and PLAYER_OVERALL_MAX == 200


def roto_salary_points(value: object) -> float:
    """Score cheaper salaries higher; missing/invalid values stay missing."""
    salary = valid_number_or_nan(value)
    if pd.isna(salary) or salary < 0:
        return float("nan")
    return float(next((points for max_salary, points in ROTO_SALARY_TIERS
                       if salary <= max_salary), 0))


def roto_projection_points(source: str, value: object) -> float:
    """Return highest earned tier or NaN when projection data is missing.

    A genuine reported numeric zero scores 0, but absent data is never
    converted into a fabricated zero. Salary is inverse to other metrics.
    """
    if source == "Player Price":
        return roto_salary_points(value)
    number = valid_number_or_nan(value)
    if pd.isna(number) or number < 0:
        return float("nan")
    tiers = ROTO_PROJECTION_TIERS[source]
    return float(next((points for cutoff, points in reversed(tiers) if number >= cutoff), 0))


def valid_number_or_nan(value: object) -> float:
    """Convert numeric source input; blanks, errors and infinities stay unavailable.

    Crucially, a genuine zero remains 0.0. Invalid source values are not replaced
    by made-up numbers and cannot poison a player's total.
    """
    try:
        number = float(value)
    except (ValueError, TypeError, OverflowError):
        return float("nan")
    return number if math.isfinite(number) else float("nan")


def player_raw_stat_values(entry: dict, boxes: list[dict]) -> tuple[dict, list[str]]:
    """Flatten the nine player groups, checking duplicate column names."""
    result = {}
    conflicts = []
    for box, group in zip(boxes, entry["groups"]):
        for label, raw in zip(box["fields"], group):
            number = valid_number_or_nan(raw)
            if label in result:
                existing = result[label]
                if pd.notna(existing) and pd.notna(number) and not math.isclose(
                    float(existing), float(number), abs_tol=1e-8
                ):
                    conflicts.append(label)
                    result[label] = float("nan")
                elif pd.isna(existing) and pd.notna(number) and label not in conflicts:
                    result[label] = number
            else:
                result[label] = number
    # Additional headers after BD (the original 56-column, 9 x 6 layout) are
    # real source stats, not a tenth six-stat group. In the Nov 22 workbook,
    # column BE (57) is FTM. Keep it available to charts and player detail.
    for label, raw in entry.get("_extra_stats", {}).items():
        number = valid_number_or_nan(raw)
        if label in result:
            existing = result[label]
            if pd.notna(existing) and pd.notna(number) and not math.isclose(
                float(existing), float(number), abs_tol=1e-8
            ):
                conflicts.append(label)
                result[label] = float("nan")
            elif pd.isna(existing) and pd.notna(number) and label not in conflicts:
                result[label] = number
        else:
            result[label] = number

    # There is no standalone player 3PM in the verified Nov 22 workbook.
    # Its total is corner three makes + above-the-break three makes. Do not
    # confuse attempt columns (C3PA/ATB 3PA) with made three-pointers.
    # Prefer a real 3PM source if future workbooks provide one.
    if "3PM" not in result or pd.isna(result["3PM"]):
        corner = result.get("C3PM", float("nan"))
        above_break = result.get("ATB 3PM", float("nan"))
        if pd.notna(corner) and pd.notna(above_break):
            result["3PM"] = round(float(corner) + float(above_break), 1)

    # Recover total field goals from 2PM + total 3PM when no FGM exists.
    if "FGM" not in result or pd.isna(result["FGM"]):
        two = result.get("2PM", float("nan"))
        three = result.get("3PM", float("nan"))
        if pd.notna(two) and pd.notna(three):
            result["FGM"] = round(float(two) + float(three), 1)
    return result, conflicts


# PLAYER STATS VIEWER — seven groups of eight metrics.
# Historical averages come only from the Player Stats workbook.
# DFS / player prop projections come only from the selected projections CSV.
# Lineup appearances come only from the optimized lineup CSV (raw counts, not Roto points).
# Props prefer ESPN Game Total and Expected Team Total from Team Stats;
# older workbooks may fall back to Expected Offense when betting lines are absent.
# Missing fields stay blank; a real reported zero stays zero.
PLAYER_CHART_GROUPS = {
    "Interior": [
        ("RA FGM", ("RA FGM",)),
        ("RA FGA", ("RA FGA",)),
        ("NRA FGM", ("NRA FGM",)),
        ("NRA FGA", ("NRA FGA",)),
        ("FTM", ("FTM",)),
        ("FTA", ("FTA",)),
        ("%PTS Paint", ("%PTS PITP", "%PTS PAINT")),
        ("%PTS FT", ("%PTS FT",)),
    ],
    "Exterior": [
        ("MR FGM", ("MR FGM",)),
        ("MR FGA", ("MR FGA",)),
        ("C 3PTM", ("C3PM", "C3PTM")),
        ("C 3PTA", ("C3PA", "C3PTA")),
        ("ATB 3PTM", ("ATB 3PM", "ATB 3PTM")),
        ("ATB 3PTA", ("ATB 3PA", "ATB 3PTA")),
        ("%PTS MR", ("%PTS 2PT MR", "%PTS MR")),
        ("%PTS 3PT", ("%PTS 3PT", "%PTS 3Pointers")),
    ],
    "Tendency": [
        ("%FGA 2PT", ("%FGA 2PT",)),
        ("%FGA 3PT", ("%FGA 3PT",)),
        ("2FGM %AST", ("2FGM %AST",)),
        ("2FGM %UAST", ("2FGM %UAST",)),
        ("3FGM %AST", ("3FGM %AST",)),
        ("3FGM %UAST", ("3FGM %UAST",)),
        ("%PTS OffTO", ("%PTS OffTO",)),
        ("%PTS FBPs", ("%PTS FBPs",)),
    ],
    "Production": [
        ("ORPG", ("OREB", "ORPG")),
        ("OREB%", ("%OREB", "OREB%")),
        ("DRPG", ("DREB", "DRPG")),
        ("DREB%", ("%DREB", "DREB%")),
        ("APG", ("AST", "APG")),
        ("AST%", ("%AST", "AST%")),
        ("Double Double", ("DD2", "Double Double")),
        ("Triple Double", ("TD3", "Triple Double")),
    ],
    "Fantasy": [
        ("SPG", ("STL", "SPG")),
        ("STL%", ("%STL", "STL%")),
        ("BPG", ("BLK", "BPG")),
        ("BLK%", ("%BLK", "BLK%")),
        ("TOPG", ("TOV", "TOPG")),
        ("TOV%", ("%TOV", "TOV%")),
        ("PF", ("PF",)),
        ("PF%", ("%PF", "PF%")),
    ],
    "DFS": [
        ("Price", ("Player Price",)),
        ("MPG", ("MPG", "MIN")),
        ("Projected Minutes", ("Projected Minutes",)),
        ("Projected Fantasy", ("Projected Fantasy Points",)),
        ("Floor", ("Projected Floor",)),
        ("Ceiling", ("Projected Ceiling",)),
        ("PPD", ("Price per Dollar",)),
        ("# of Lineups", ("Lineup Appearances",)),
    ],
    "Props": [
        ("PPG", ("PPG",)),
        ("Projected Points", ("Projected Points",)),
        ("RPG", ("RPG",)),
        ("Projected Rebounds", ("Projected Rebounds",)),
        ("APG", ("AST", "APG")),
        ("Projected Assists", ("Projected Assists",)),
        ("Expected Game Total", ("Expected Game Total",)),
        ("Expected Team Total", ("Expected Team Total",)),
    ],
}
assert len(PLAYER_CHART_GROUPS) == 7
assert all(len(stats) == 8 for stats in PLAYER_CHART_GROUPS.values())


def _player_header_key(header: str) -> str:
    """Normalize only spelling; preserve the distinction between % and raw count."""
    normalized = str(header).strip().casefold().replace("percentage", "pct")
    normalized = normalized.replace("percent", "pct").replace("%", "pct")
    return re.sub(r"[^a-z0-9]", "", normalized)


def resolved_player_source(raw_values: dict, aliases: tuple[str, ...]) -> str | None:
    """Find the unique workbook column corresponding to one approved header alias."""
    exact_index = {}
    normalized_index = {}
    for header in raw_values:
        exact_index.setdefault(str(header).strip().casefold(), []).append(header)
        normalized_index.setdefault(_player_header_key(header), []).append(header)
    for alias in aliases:
        matches = exact_index.get(str(alias).strip().casefold(), [])
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            return None
    for alias in aliases:
        matches = normalized_index.get(_player_header_key(alias), [])
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            return None
    return None


def expected_team_total_from_workbook(team_extra: dict) -> float:
    """Prefer the ESPN market-implied team total; fall back for older workbooks.

    AJ is the workbook's Expected Team Total. If its formula cache is empty,
    calculate the equivalent using the same workbook's AH/AI inputs.
    Only legacy workbooks without odds fields use AB Expected Offense.
    """
    supplied = valid_number_or_nan(team_extra.get("Expected Team Total"))
    if pd.notna(supplied):
        return supplied
    total = valid_number_or_nan(team_extra.get("Game Total"))
    spread = valid_number_or_nan(team_extra.get("Spread"))
    if pd.notna(total) and pd.notna(spread):
        return round((total - spread) / 2.0, 2)
    return valid_number_or_nan(team_extra.get("Expected Offense"))


def expected_game_total_from_workbook(team_extra: dict, opponent_extra: dict) -> float:
    """Use the ESPN game total when present; otherwise combine available estimates."""
    total = valid_number_or_nan(team_extra.get("Game Total"))
    if pd.notna(total):
        return total
    team_value = expected_team_total_from_workbook(team_extra)
    opponent_value = expected_team_total_from_workbook(opponent_extra)
    if pd.notna(team_value) and pd.notna(opponent_value):
        return round(team_value + opponent_value, 2)
    return float("nan")


def espn_game_odds_table(team_data: dict, teams: list[str]) -> pd.DataFrame:
    """One team-perspective row per included club, using uploaded workbook odds only."""
    extras = {entry["name"]: entry.get("_team_projection", {})
              for entry in team_data["entries"]}
    rows = []
    for team in dict.fromkeys(teams):
        extra = extras.get(team, {})
        if not any(k in extra for k in ("Game Total", "Spread", "Expected Team Total")):
            continue  # An older workbook with no ESPN odds section.
        opponent_name = str(extra.get("Opponent") or "").strip()
        opponent_data = extras.get(opponent_name, {})
        implied_team = expected_team_total_from_workbook(extra)
        game_total = expected_game_total_from_workbook(extra, opponent_data)
        opponent_implied = expected_team_total_from_workbook(opponent_data)
        if pd.isna(opponent_implied) and pd.notna(game_total) and pd.notna(implied_team):
            opponent_implied = round(game_total - implied_team, 2)
        rows.append({
            "Team": team,
            "Opponent": opponent_name,
            "Game Time (PT)": extra.get("Game Time (PT)", ""),
            "Game Total": game_total,
            "Spread": valid_number_or_nan(extra.get("Spread")),
            "Expected Team Total": implied_team,
            "Expected Opponent Total": opponent_implied,
            "Venue": extra.get("Venue", ""),
        })
    return pd.DataFrame(rows, columns=[
        "Team", "Opponent", "Game Time (PT)", "Game Total", "Spread",
        "Expected Team Total", "Expected Opponent Total", "Venue",
    ])


def player_chart_table(player_data: dict, chart_name: str, metadata: pd.DataFrame,
                       team_data: dict, selected_pairs: list[tuple[str, str]]) -> tuple[pd.DataFrame, list[str]]:
    """Build a requested 8-metric Player Stats view without changing scoring data."""
    specs = PLAYER_CHART_GROUPS[chart_name]
    box_fields = [field for box in player_data["boxes"] for field in box["fields"]]
    source_fields = list(dict.fromkeys(box_fields + player_data.get("extra_fields", [])))
    probe = {field: None for field in source_fields}
    if metadata.empty:
        meta_index = {}
    else:
        meta_index = {(r["Player"], r["Team"]): r
                      for r in metadata.to_dict("records")}
    team_expectations = {t["name"]: t.get("_team_projection", {})
                         for t in team_data["entries"]}
    code_to_name = {TEAM_CODES[name]: name for name in team_expectations if name in TEAM_CODES}
    paired = {a: b for a, b in selected_pairs}
    paired.update({b: a for a, b in selected_pairs})
    def stat_value(raw, *aliases):
        source = resolved_player_source(probe, aliases)
        return valid_number_or_nan(raw.get(source)) if source else float("nan")

    rows = []
    for entry in player_data["entries"]:
        key = (entry["name"], entry["team"])
        meta = meta_index.get(key, {})
        raw, _conflicts = player_raw_stat_values(entry, player_data["boxes"])
        team = code_to_name.get(entry["team"])
        team_extra = team_expectations.get(team, {})
        # The selected game pairing is authoritative. Fall back to the workbook
        # "Opponent" text, then the matched projections OPP team only when valid.
        opponent = paired.get(team)
        if not opponent:
            named = team_extra.get("Opponent")
            if named in team_expectations and named != team:
                opponent = named
        if not opponent:
            candidate = code_to_name.get(entry.get("_projection_opponent_code"))
            if candidate and candidate != team:
                opponent = candidate
        # New workbooks carry ESPN market totals in AH/AJ, while older ones
        # only carry statistical Expected Offense estimates in AB.
        opponent_extra = team_expectations.get(opponent, {})
        expected_team = expected_team_total_from_workbook(team_extra)
        expected_game = expected_game_total_from_workbook(team_extra, opponent_extra)
        record = {"Player": entry["name"], "Team": entry["team"]}
        for label, aliases in specs:
            if chart_name == "DFS":
                if label == "MPG":
                    val = stat_value(raw, *aliases)
                else:
                    val = valid_number_or_nan(meta.get(aliases[0]))
            elif chart_name == "Props":
                if label in ("Projected Points", "Projected Rebounds", "Projected Assists"):
                    val = valid_number_or_nan(meta.get(label))
                elif label == "Expected Team Total":
                    val = expected_team
                elif label == "Expected Game Total":
                    val = expected_game
                elif label == "RPG":
                    offensive = stat_value(raw, "OREB")
                    defensive = stat_value(raw, "DREB")
                    val = offensive + defensive if pd.notna(offensive) and pd.notna(defensive) else float("nan")
                else:
                    val = stat_value(raw, *aliases)
            elif chart_name == "Tendency" and label == "3FGM %AST":
                val = stat_value(raw, *aliases)
                if pd.isna(val):
                    unassisted = stat_value(raw, "3FGM %UAST")
                    threes = valid_number_or_nan(raw.get("3PM"))
                    # The percentage is undefined for players with no made 3s.
                    if pd.notna(unassisted) and 0 <= unassisted <= 100 and pd.notna(threes) and threes > 0:
                        val = round(100 - unassisted, 1)
            else:
                val = stat_value(raw, *aliases)
            record[label] = val
        rows.append(record)

    cols = ["Player", "Team", *[label for label, _ in specs]]
    result = pd.DataFrame(rows, columns=cols)
    missing = [c for c in cols[2:] if result[c].notna().sum() == 0]
    return result, missing


def show_player_chart(frame: pd.DataFrame, missing: list[str], chart_name: str) -> None:
    """Sort/export one eight-metric player view."""
    if frame.empty:
        st.info("No eligible players in this chart.")
        return
    if missing:
        st.caption("Not found in this workbook (shown blank): " + ", ".join(missing))
    numeric = list(frame.columns[2:])
    if numeric:
        left, right = st.columns([2, 1])
        with left:
            selected = st.selectbox("Sort by", numeric + ["Player", "Team"],
                                    key="player_chart_sort_" + csv_col_key(chart_name))
        with right:
            descending = st.selectbox("Order", ["Highest first", "Lowest first"],
                                      key="player_chart_direction_" + csv_col_key(chart_name))
        frame = frame.sort_values(selected, ascending=descending == "Lowest first",
                                  kind="stable", na_position="last")
    st.dataframe(frame, hide_index=True, use_container_width=True,
                 height=min(730, 85 + 34 * len(frame)),
                 column_config=one_decimal_config(frame))
    st.download_button("Download this chart CSV", frame.to_csv(index=False).encode("utf-8-sig"),
                       file_name="orange_nba_" + csv_col_key(chart_name) + ".csv",
                       mime="text/csv", key="player_chart_download_" + csv_col_key(chart_name))


def build_player_stat_rows(
    player_data: dict, team_data: dict, team_to_opponent: dict[str, str],
) -> tuple[pd.DataFrame, list[str]]:
    """One original per-game value per player/stat; opponent used for bonuses.

    Every projection-eligible player is retained, even when their opponent is
    unavailable. Invalid or conflicting source values are left missing, while
    other valid statistics remain usable for partial scores.
    """
    team_by_code = {TEAM_CODES[name]: name for name in team_data if name in TEAM_CODES}
    rows = []
    warnings = set()
    for entry in player_data["entries"]:
        team_name = team_by_code.get(entry["team"])
        opponent = team_to_opponent.get(team_name, "") if team_name else ""
        # If ESPN/manual slate matching cannot provide an opponent, use the
        # projection CSV's OPP field only when it resolves to a team in this
        # workbook. Never assume the player's own team is the opponent.
        if not opponent:
            fallback = team_by_code.get(entry.get("_projection_opponent_code", ""), "")
            if fallback and fallback != team_name:
                opponent = fallback
        player_values, conflicts = player_raw_stat_values(entry, player_data["boxes"])
        if conflicts:
            warnings.add(f'{entry["name"]} ({entry["team"]}): conflicting values for '
                         + ", ".join(conflicts))
        for stat in PLAYER_DETAIL_STATS:
            rows.append({
                "Player": entry["name"], "Team": entry["team"],
                "Opponent": opponent, "Stat": stat,
                "Value": valid_number_or_nan(player_values.get(stat)),
            })
    columns = ["Player", "Team", "Opponent", "Stat", "Value"]
    return pd.DataFrame(rows, columns=columns), sorted(warnings)


# Player ranking and CSV handling are defined below the workbook display helpers.

# The scoreboard provides MATCHUPS ONLY. Stats are always read from the uploaded workbook.
# The service is not affiliated with this app and can occasionally be unavailable.
SCHEDULE_API = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/scoreboard"
ESPN_CODE_ALIASES = {
    "GS": "GSW", "SA": "SAS", "NY": "NYK", "NO": "NOP", "PHO": "PHX",
    "UTAH": "UTA", "WSH": "WAS", "NJ": "BKN",
}


def default_slate_date() -> date:
    """Today in U.S. Eastern time; users can select any other game date.

    Uploaded filenames never determine, limit, or change the slate date.
    """
    return datetime.now(ZoneInfo("America/New_York")).date()


@st.cache_data(ttl=900, show_spinner=False)
def fetch_nba_slate(slate_date: date) -> list[dict]:
    """Fetch all NBA games for one US game date, including a full 15-game slate."""
    url = f"{SCHEDULE_API}?dates={slate_date:%Y%m%d}&limit=100"
    request = Request(url, headers={
        "Accept": "application/json", "User-Agent": "OrangeNBAViewer/1.0",
    })
    try:
        with urlopen(request, timeout=12) as response:
            payload = json.load(response)
    except (OSError, HTTPError, URLError, ValueError) as exc:
        raise RuntimeError(f"Schedule lookup failed: {exc}") from exc
    if not isinstance(payload, dict) or not isinstance(payload.get("events"), list):
        raise RuntimeError("Schedule service returned an unexpected response.")
    games = []
    for event in payload["events"]:
        competitions = event.get("competitions") or []
        if not competitions:
            continue
        teams = competitions[0].get("competitors") or []
        if len(teams) != 2:
            continue
        ordered = sorted(teams, key=lambda x: 0 if x.get("homeAway") == "away" else 1)
        participants = []
        for item in ordered:
            raw_team = item.get("team") or {}
            participants.append({
                "name": str(raw_team.get("displayName") or "").strip(),
                "code": str(raw_team.get("abbreviation") or "").upper().strip(),
            })
        games.append({"away": participants[0], "home": participants[1],
                      "game_id": str(event.get("id", "")),
                      "status": str(((event.get("status") or {}).get("type") or {}).get("description", ""))})
    return games


def map_schedule_team(info: dict, workbook_teams: list[str]) -> str | None:
    """Use exact workbook names when possible; otherwise normalize NBA abbreviations."""
    original_name = (info.get("name") or "").strip()
    name_match = {name.casefold(): name for name in workbook_teams}
    if original_name.casefold() in name_match:
        return name_match[original_name.casefold()]
    code = ESPN_CODE_ALIASES.get((info.get("code") or "").upper(),
                                (info.get("code") or "").upper())
    by_code = {TEAM_CODES[name]: name for name in workbook_teams if name in TEAM_CODES}
    if code in by_code:
        return by_code[code]
    # Display name aliases (e.g. LA Clippers vs Los Angeles Clippers).
    team_code = TEAM_CODES.get(original_name)
    return by_code.get(team_code) if team_code else None


def match_schedule_to_workbook(
    schedule_games: list[dict], workbook_teams: list[str],
) -> tuple[list[tuple[str, str]], list[str]]:
    """Return valid away/home pairs plus unrecognized/unavailable games."""
    pairs = []
    unavailable = []
    seen = set()
    for game in schedule_games:
        away = map_schedule_team(game["away"], workbook_teams)
        home = map_schedule_team(game["home"], workbook_teams)
        if not away or not home:
            unavailable.append(f'{game["away"]["name"]} at {game["home"]["name"]}')
            continue
        if away == home or away in seen or home in seen:
            unavailable.append(f"Conflicting schedule pairing: {away} at {home}")
            continue
        seen.update((away, home))
        pairs.append((away, home))
    return pairs, unavailable


def validate_matchup_pairs(
    pairs: list[tuple[str, str]], workbook_teams: list[str],
) -> tuple[list[tuple[str, str]], list[str]]:
    """Never assign two defenses to a team or silently invent missing opponents."""
    recognized = set(workbook_teams)
    seen = set()
    valid = []
    warnings = []
    for away, home in pairs:
        if not away or not home:
            continue
        if away not in recognized or home not in recognized:
            warnings.append(f"Team not in workbook: {away} at {home}")
        elif away == home:
            warnings.append(f"A team cannot play itself: {away}")
        elif away in seen or home in seen:
            warnings.append(f"Duplicate team assigned to more than one game: {away} at {home}")
        else:
            valid.append((away, home))
            seen.update((away, home))
    return valid, warnings


def game_label(away: str, home: str) -> str:
    return f"{away} at {home}"


def matchup_selection_ui(teams: list[str]) -> list[tuple[str, str]]:
    """Pick the date once for every tab, then auto-pair the complete slate."""
    st.subheader("Game slate")
    st.caption("Game pairings are fetched from ESPN's NBA daily scoreboard feed. All statistics "
               "and player calculations still come exclusively from your uploaded Excel file.")
    col_date, col_mode, col_refresh = st.columns([1.1, 1.8, 0.65])
    with col_date:
        slate_date = st.date_input("Slate date (U.S. Eastern)",
                                   value=default_slate_date(),
                                   key="selected_slate_date")
    with col_mode:
        source_mode = st.radio("Find matchups", ["Automatic daily schedule", "Enter / edit matchups"],
                               horizontal=True, key="schedule_source_mode")
    with col_refresh:
        st.write(" ")
        if st.button("Refresh", help="Reload the schedule from its provider"):
            fetch_nba_slate.clear()

    raw_games = []
    api_error = None
    try:
        raw_games = fetch_nba_slate(slate_date)
    except RuntimeError as exc:
        api_error = str(exc)
    available, unmatched = match_schedule_to_workbook(raw_games, teams)

    if source_mode == "Automatic daily schedule" and api_error:
        st.warning(f"{api_error} Switch to 'Enter / edit matchups' below to continue.")
        source_mode = "Enter / edit matchups"
    elif source_mode == "Automatic daily schedule" and not raw_games:
        st.warning(f"No NBA games returned for {slate_date:%B %d, %Y}. "
                   "Check the selected date, or enter matchups manually.")

    if source_mode == "Enter / edit matchups":
        st.caption("Choose each game's two teams. Add rows for more games, and leave unused "
                   "rows blank. No matches are assumed if the schedule is unavailable.")
        starting_rows = ([{"Away": a, "Home": b} for a, b in available]
                         if available else [{"Away": None, "Home": None}])
        with st.form("manual_games_form"):
            edited = st.data_editor(
                pd.DataFrame(starting_rows, columns=["Away", "Home"]),
                hide_index=True, num_rows="dynamic", use_container_width=True,
                column_config={
                    "Away": st.column_config.SelectboxColumn("Away", options=sorted(teams)),
                    "Home": st.column_config.SelectboxColumn("Home", options=sorted(teams)),
                }, key=f"manual_games_{slate_date:%Y%m%d}",
            )
            applied = st.form_submit_button("Use these matchups")
        if applied:
            st.session_state["manual_pairs"] = [
                (str(row.Away), str(row.Home))
                for row in edited.itertuples(index=False)
                if pd.notna(row.Away) and pd.notna(row.Home)
            ]
            st.session_state["manual_date"] = slate_date.isoformat()
        pairs = (st.session_state.get("manual_pairs", [])
                 if st.session_state.get("manual_date") == slate_date.isoformat() else [])
        if not pairs:
            st.info("Enter the games above and click 'Use these matchups' to show opponent labels; player scores work without matchups.")
    else:
        pairs = available

    pairs, invalid = validate_matchup_pairs(pairs, teams)
    for problem in invalid:
        st.warning(problem)
    if unmatched and source_mode == "Automatic daily schedule":
        st.warning(f"{len(unmatched)} scheduled games were excluded because a team was absent "
                   "from the workbook or a team appeared twice.")
        with st.expander("See excluded games"):
            st.write(unmatched)

    if pairs:
        labels = [game_label(a, b) for a, b in pairs]
        selected_labels = st.multiselect("Games to include (all selected by default)",
                                         labels, default=labels,
                                         key=f"games_to_include_{slate_date:%Y%m%d}_{source_mode}")
        pairs = [pair for pair, label in zip(pairs, labels) if label in selected_labels]
    if pairs:
        st.caption(f"{len(pairs)} games · {len(pairs) * 2} teams mapped to opponent defenses")
        with st.expander("See selected games"):
            st.dataframe(pd.DataFrame([{"Away": a, "Home": b} for a, b in pairs]),
                         hide_index=True, use_container_width=True)
    else:
        st.warning("No matchups selected. Team and player source tables still work; "
                   "player rankings still work without opponent adjustments.")
    return pairs

XML_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XML_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XML_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _col_idx(reference: str) -> int:
    match = re.match(r"[A-Z]+", reference)
    if not match:
        raise ValueError(f"Invalid Excel cell: {reference}")
    index = 0
    for char in match.group():
        index = index * 26 + ord(char) - ord("A") + 1
    return index - 1


def _read_rows(archive: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> list[list]:
    root = ET.fromstring(archive.read(sheet_path))
    rows = []
    for row in root.findall(f".//{XML_MAIN}sheetData/{XML_MAIN}row"):
        values = {}
        for cell in row.findall(f"{XML_MAIN}c"):
            position = _col_idx(cell.attrib.get("r", "A1"))
            kind = cell.attrib.get("t", "")
            raw = cell.findtext(f"{XML_MAIN}v")
            if kind == "inlineStr":
                value = "".join(tag.text or "" for tag in cell.findall(f".//{XML_MAIN}t"))
            elif raw is None:
                value = None
            elif kind == "s":
                value = shared[int(raw)]
            elif kind in ("str", "e"):
                value = raw
            elif kind == "b":
                value = raw == "1"
            else:
                try:
                    number = float(raw)
                    value = int(number) if number.is_integer() else number
                except ValueError:
                    value = raw
            values[position] = value
        if values:
            rows.append([values.get(i) for i in range(max(values) + 1)])
    return rows


def parse_workbook(contents: bytes) -> dict:
    """Read the two-tab master workbook from upload bytes (no local file needed)."""
    with zipfile.ZipFile(io.BytesIO(contents)) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = [
                "".join(t.text or "" for t in si.findall(f".//{XML_MAIN}t"))
                for si in root.findall(f"{XML_MAIN}si")
            ]

        book = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relpaths = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall(f"{XML_PKG}Relationship")}
        source_sheets = {}
        for node in book.findall(f".//{XML_MAIN}sheets/{XML_MAIN}sheet"):
            target = relpaths[node.attrib[XML_REL + "id"]]
            if target.startswith("/"):
                filename = target.lstrip("/")
            elif target.startswith("xl/"):
                filename = target
            else:
                filename = "xl/" + target
            source_sheets[node.attrib["name"]] = _read_rows(archive, filename, shared)

    required = {"Team Stats": (25, 1, 4), "Player Stats": (56, 2, 9)}
    data = {}
    for sheet, (width, stat_start, expected_boxes) in required.items():
        if sheet not in source_sheets:
            raise ValueError(f'Missing required tab "{sheet}". Use an NBA master-format workbook.')
        rows = source_sheets[sheet]
        if len(rows) < 3:
            raise ValueError(f'"{sheet}" needs a grouped-header row, a column-header row, and data.')
        headers = (rows[1] + [None] * width)[:width]
        if any(not isinstance(h, str) or not h.strip() for h in headers):
            raise ValueError(f'"{sheet}" must have all {width} column headers in row 2.')
        if len(headers) != width or (width - stat_start) != 6 * expected_boxes:
            raise ValueError(f'Unexpected column layout in "{sheet}".')
        # Supplemental player columns may follow the original 56. Example:
        # column BE is FTM in the Nov 22 workbook. Preserve every named trailing
        # header without changing the nine original six-stat box boundaries.
        extra_fields = []
        team_extra_fields = []
        if sheet == "Team Stats":
            for extra_header in rows[1][width:]:
                if isinstance(extra_header, str) and extra_header.strip():
                    team_extra_fields.append(extra_header.strip())
                else:
                    break
        if sheet == "Player Stats":
            for extra_header in rows[1][width:]:
                if isinstance(extra_header, str) and extra_header.strip():
                    extra_fields.append(extra_header.strip())
                else:
                    break
            if len(set(headers + extra_fields)) != len(headers + extra_fields):
                raise ValueError('Player Stats contains duplicate source headers.')

        boxes = [
            {"name": (f"Team Box #{i+1}" if sheet == "Team Stats" else f"Player Stats #{i+1}"),
             "fields": headers[stat_start+6*i:stat_start+6*(i+1)]}
            for i in range(expected_boxes)
        ]
        entries = []
        for source in rows[2:]:
            data_width = width + len(extra_fields)
            row = (source + [None] * data_width)[:data_width]
            name = str(row[0]).strip() if row[0] is not None else ""
            if not name:
                continue  # Empty separator rows are not players or teams.

            # Uploaded daily workbooks may include supplemental tables BELOW the
            # standard 25-column team grid, plus source notes below Player Stats.
            # _read_rows intentionally omits completely blank Excel rows, so a
            # blank-row delimiter alone is not enough to locate these sections.
            # Stop at their header/title instead of counting those rows as
            # additional teams/players. The full workbook is left untouched.
            upper_name = name.upper()
            is_new_section = (
                name.casefold() == str(headers[0]).strip().casefold()
                or upper_name.startswith((
                    "MATCHUP CALCULATION INPUTS", "MATCHUP SOURCE TABLE",
                    "MATCHUP PROJECTION INPUTS", "MATCHUP PROJECTIONS",
                    "SOURCE:", "NOTES:", "DATA SOURCE:",
                ))
            )
            if sheet == "Player Stats" and (
                row[1] is None or str(row[1]).strip() == ""
            ):
                is_new_section = True  # Footer note, not an NBA player.
            if is_new_section:
                if entries:
                    break
                continue

            entry = {"name": name,
                     "groups": [row[stat_start+6*i:stat_start+6*(i+1)] for i in range(expected_boxes)]}
            if sheet == "Player Stats":
                entry["team"] = str(row[1]).strip()
                entry["_extra_stats"] = dict(zip(extra_fields, row[width:data_width]))
            elif team_extra_fields:
                entry["_team_projection"] = dict(zip(
                    team_extra_fields, (source + [None] * (width + len(team_extra_fields)))[
                        width:width + len(team_extra_fields)]))
            entries.append(entry)
        if not entries:
            raise ValueError(f'"{sheet}" has no data rows.')
        data[sheet] = {"boxes": boxes, "entries": entries,
                       "extra_fields": extra_fields}

    team_list = [t["name"] for t in data["Team Stats"]["entries"]]
    if len(team_list) != len(set(team_list)):
        raise ValueError("Team Stats contains duplicate team names.")
    data["teamNames"] = team_list
    return data


@st.cache_data(show_spinner=False)
def cached_workbook(contents: bytes) -> dict:
    return parse_workbook(contents)


def six_stat_table(sheet_data: dict, sheet_name: str, box_index: int) -> pd.DataFrame:
    """Make one sortable table containing identifiers and ONLY one six-stat box."""
    box = sheet_data["boxes"][box_index]
    fields = box["fields"]
    if len(fields) != 6 or len(set(fields)) != 6:
        raise ValueError(f'{box["name"]} needs six distinct statistic headers.')

    is_player = sheet_name == "Player Stats"
    rows = []
    for entry in sheet_data["entries"]:
        record = {"Player": entry["name"], "Team": entry["team"]} if is_player else {"Team": entry["name"]}
        for field, value in zip(fields, entry["groups"][box_index]):
            record[field] = value
        rows.append(record)

    identifier_cols = ["Player", "Team"] if is_player else ["Team"]
    # Preserve headers even when the projections-only filter leaves no players.
    frame = pd.DataFrame(rows, columns=identifier_cols + fields)
    for field in fields:
        # Excel blank cells and '-' are missing values, NOT zero.
        frame[field] = pd.to_numeric(frame[field], errors="coerce")
    return frame


def show_six_stat_tab(frame: pd.DataFrame, *, section: str, box_name: str) -> None:
    """Show six stats per tab with explicit numeric sorting and CSV export."""
    if frame.empty:
        st.info("No matching rows. Change the filters to show results.")
        return

    identifier_columns = 2 if section == "players" else 1
    stat_fields = list(frame.columns[identifier_columns:])
    if len(stat_fields) != 6:
        raise ValueError(f"Expected exactly 6 stats in {box_name}.")

    col_sort, col_order = st.columns([2, 1])
    with col_sort:
        sort_field = st.selectbox("Sort by", stat_fields + list(frame.columns[:identifier_columns]),
                                  key=f"{section}_{box_name}_sort")
    with col_order:
        direction = st.selectbox("Order", ["Highest first", "Lowest first"],
                                 index=0 if sort_field in stat_fields else 1,
                                 key=f"{section}_{box_name}_direction")

    sorted_frame = frame.sort_values(sort_field, ascending=(direction == "Lowest first"),
                                     kind="stable", na_position="last").reset_index(drop=True)
    st.caption(f"{len(sorted_frame)} rows · 6 statistics · Click any column heading to sort, "
               "or use the controls above.")
    if section == "teams":
        # Pandas Styler highlights each numeric CELL, not entire rows or columns.
        # The underlying DataFrame remains numeric and sortable/exportable.
        styled = sorted_frame.style
        for field in stat_fields:
            styled = styled.map(lambda value, col=field: style_team_cell(value, col), subset=[field])
        formatters = {field: "{:.1f}" for field in stat_fields}
        styled = styled.format(formatters, na_rep="—")
        st.dataframe(styled, hide_index=True, use_container_width=True,
                     height=min(640, 80 + 35 * len(sorted_frame)))
        with st.expander("View thresholds for this box"):
            threshold_df = threshold_reference(stat_fields)
            st.dataframe(threshold_df, hide_index=True,
                         use_container_width=True,
                         column_config=one_decimal_config(threshold_df))
            st.caption("Tier boundaries apply to the raw numeric value. Explicit "
                       "'or lower' limits include that exact number; values just above "
                       "it are neutral until the favorable threshold.")
    else:
        st.dataframe(sorted_frame, hide_index=True, use_container_width=True,
                     height=min(640, 80 + 35 * len(sorted_frame)),
                     column_config=one_decimal_config(sorted_frame))
    st.download_button("Download this group's CSV", sorted_frame.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"orange_nba_{section}_{box_name.lower().replace(' ', '_').replace('#','')}.csv",
                       mime="text/csv", key=f"{section}_{box_name}_csv")


# OPTIONAL LINEUP AND PROJECTION CSV IMPORTS
# Player matching is exact after normalization; ambiguous names are skipped.
import unicodedata

PROJECTION_LABELS = [
    "Player Price", "Projected Floor", "Projected Ceiling",
    "Price per Dollar", "Smash", "Projected Fantasy Points",
    "Projected Minutes",
]
# These three raw projections populate Props only; no change to Roto weights.
PROP_PROJECTION_LABELS = ["Projected Points", "Projected Rebounds", "Projected Assists"]
PROJECTION_SYNONYMS = {
    "Player Price": ("player price", "salary", "price", "cost"),
    "Projected Floor": ("projected floor", "proj floor", "floor", "floor points"),
    "Projected Ceiling": ("projected ceiling", "proj ceiling", "ceiling", "ceil"),
    "Price per Dollar": ("price per dollar", "price/dollar", "price per $", "ppd", "value", "p/$", "pts per dollar", "fpts/$", "fpts per dollar"),
    "Smash": ("smash", "smash %", "smash rate", "smash probability"),
    "Projected Fantasy Points": ("fpts", "projected fantasy points", "projected fpts", "proj fpts", "fantasy points projection", "projected points"),
    "Projected Minutes": ("minutes", "projected minutes", "proj minutes", "projected min", "proj min", "mins", "min"),
    "Projected Points": ("pts", "projected pts", "proj pts", "points projection"),
    "Projected Rebounds": ("reb", "projected reb", "proj reb", "projected rebounds"),
    "Projected Assists": ("ast", "projected ast", "proj ast", "projected assists"),
}
NAME_SYNONYMS = ("player", "player name", "name", "full name", "nickname", "athlete", "name + id", "player + id")
TEAM_SYNONYMS = ("team", "team abbreviation", "team abbrev", "teamabbr", "tm")
COUNT_SYNONYMS = ("appearances", "lineup count", "lineups", "lineup appearances", "count", "frequency", "times selected")
LINEUP_ID_SYNONYMS = ("lineup id", "lineup_id", "lineup", "entry id", "entry_id", "entry")

# Known name variants in DraftKings/optimizer CSVs versus NBA workbook names.
# These are *only* fallbacks after an exact normalized-name lookup fails.
# Team matching is still required whenever the CSV supplies a team.
# Keys and values use clean_player_name() normalization (lowercase, no punctuation).
PLAYER_NAME_ALIASES = {
    "alexandresarr": "alexsarr",
    "carltoncarrington": "bubcarrington",
    "craigporter": "craigporterjr",
    "dariusacuff": "dariusacuffjr",
    "gregoryjackson": "ggjackson",
    "jaimejaquez": "jaimejaquezjr",
    "labaronphilon": "labaronphilonjr",
    "ludort": "luguentzdort",
    "morezjohnson": "morezjohnsonjr",
    "ronharper": "ronharperjr",
    "tarrisreed": "tarrisreedjr",
    "terrenceshannon": "terrenceshannonjr",
    "treymurphy": "treymurphyiii",
    "walterclayton": "walterclaytonjr",
    # DraftKings optimizer may shorten Robert Williams III in lineup cells.
    "robertwilliams": "robertwilliamsiii",
    # Nov 7 optimizer versus projections: same players, different name forms.
    "mikelbrownjr": "mikelbrown",
    "moritzwagner": "moewagner",
}


def csv_col_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def find_column(cols: list[str], choices: tuple[str, ...]) -> str | None:
    """Prefer exact CSV headers to avoid confusing FPTS with FPTS/$.

    Punctuation-insensitive matching is a fallback only, and ambiguous matches
    are rejected rather than silently attaching the wrong stat column.
    """
    exact = {str(col).strip().casefold(): col for col in cols}
    for word in choices:
        match = exact.get(str(word).strip().casefold())
        if match is not None:
            return match
    for word in choices:
        key = csv_col_key(word)
        candidates = [col for col in cols if csv_col_key(col) == key]
        # FPTS/$ is a different metric, not a spelling variant of FPTS.
        if key == "fpts":
            candidates = [col for col in candidates if "/" not in col and "$" not in col]
        if len(candidates) == 1:
            return candidates[0]
    return None


def clean_player_name(raw: object) -> str:
    if pd.isna(raw):
        return ""
    value = str(raw).strip()
    # Remove a trailing player ID, or team abbreviation inside parentheses.
    value = re.sub(r"\s*\((?:\d{2,}|[A-Z]{2,4})\)\s*$", "", value)
    value = re.sub(r"\s*\[\d{2,}\]\s*$", "", value)
    # Sometimes a lineup export is `Last, First`.
    if value.count(",") == 1 and not re.search(r"\d", value):
        last, first = (p.strip() for p in value.split(",", 1))
        if last and first:
            value = first + " " + last
    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", value.casefold())


def player_lookup_indices(players: list[dict]) -> tuple[dict, dict]:
    by_name = {}
    by_name_team = {}
    for person in players:
        key = (person["name"], person["team"])
        norm = clean_player_name(person["name"])
        if norm:
            by_name.setdefault(norm, set()).add(key)
            by_name_team.setdefault((norm, person["team"]), set()).add(key)
    return by_name, by_name_team


def normalized_team(raw: object) -> str:
    if pd.isna(raw):
        return ""
    team = str(raw).strip().upper()
    if team in TEAM_CODES.values():
        return team
    for name, code in TEAM_CODES.items():
        if name.casefold() == str(raw).strip().casefold():
            return code
    return ESPN_CODE_ALIASES.get(team, team)


def identify_player(raw: object, raw_team: object, lookup: tuple[dict, dict]):
    name = clean_player_name(raw)
    if not name:
        return None
    by_name, by_team = lookup
    team = normalized_team(raw_team)

    def candidates_for(normalized_name: str) -> set:
        return (by_team.get((normalized_name, team), set()) if team
                else by_name.get(normalized_name, set()))

    # Always prefer the athlete's exact normalized name over an alias.
    # Never override an ambiguous exact match.
    candidates = candidates_for(name)
    if len(candidates) == 1:
        return next(iter(candidates))
    if candidates:
        return None

    # Retry a specifically approved alias; never fuzzy-match unrelated names.
    canonical_name = PLAYER_NAME_ALIASES.get(name)
    if canonical_name:
        candidates = candidates_for(canonical_name)
        if len(candidates) == 1:
            return next(iter(candidates))
    return None  # Never guess which of several players was intended.


def read_optional_csv(contents: bytes) -> pd.DataFrame:
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            result = pd.read_csv(io.BytesIO(contents), sep=None, engine="python",
                                 dtype=str, encoding=encoding)
            if len(result.columns) > 0:
                result.columns = [str(c).strip() for c in result.columns]
                return result.dropna(how="all")
        except (UnicodeDecodeError, pd.errors.ParserError, ValueError):
            pass
    raise ValueError("Unable to parse the uploaded CSV. Try exporting as a standard UTF-8 CSV.")



def roster_position_columns(columns: list[str]) -> list[str]:
    positions = r"(pg|sg|sf|pf|c|g|f|util|utility|flex|guard|forward|center|cpt|captain|mvp|star)"
    found = []
    for col in columns:
        key = str(col).strip().lower()
        # Standard DFS CSV roster slots (including pandas' duplicate `.1` headers).
        if re.fullmatch(positions + r"(?:[ ._-]?\d+)?", key) or re.fullmatch(r"player[ _-]?\d+", key):
            found.append(col)
    return found


def count_lineup_appearances(source: pd.DataFrame, players: list[dict], mode: str,
                              player_col: str | None, team_col: str | None,
                              id_col: str | None, count_col: str | None,
                              roster_cols: list[str]) -> tuple[dict, list[str], int]:
    from collections import Counter
    lookup = player_lookup_indices(players)
    counts = Counter()
    unmatched = set()
    matched_appearances = 0
    if mode == "Wide lineup rows":
        if not roster_cols:
            return {}, ["No roster slot columns were selected."], 0
        for _, row in source.iterrows():
            lineup_keys = set()
            for col in roster_cols:
                raw = row[col]
                if pd.isna(raw) or not str(raw).strip():
                    continue
                key = identify_player(raw, None, lookup)
                if key is None:
                    unmatched.add(str(raw).strip())
                else:
                    lineup_keys.add(key)
            counts.update(lineup_keys)  # At most once per lineup.
            matched_appearances += len(lineup_keys)
    else:
        if not player_col:
            return {}, ["Select a player-name column."], 0
        # Prevent duplicate entries per player per lineup when a lineup ID is present.
        encountered = set()
        for row_idx, row in source.iterrows():
            raw = row[player_col]
            if pd.isna(raw) or not str(raw).strip():
                continue
            key = identify_player(raw, row[team_col] if team_col else None, lookup)
            if key is None:
                unmatched.add(str(raw).strip())
                continue
            if mode == "Precomputed count":
                value = valid_number_or_nan(row[count_col]) if count_col else float("nan")
                if pd.isna(value) or value < 0 or int(value) != value:
                    unmatched.add(f"Invalid appearance count: {raw}")
                    continue
                counts[key] += int(value)
                matched_appearances += int(value)
            else:
                lineup_id = row[id_col] if id_col and pd.notna(row[id_col]) else row_idx
                token = (key, str(lineup_id))
                if token not in encountered:
                    counts[key] += 1
                    matched_appearances += 1
                    encountered.add(token)
    return dict(counts), sorted(unmatched), matched_appearances


def csv_numeric(value: object) -> float:
    if value is None or pd.isna(value) or str(value).strip() == "":
        return float("nan")
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("$", "")
        text = text.rstrip("%xX ")
        return valid_number_or_nan(text)
    return valid_number_or_nan(value)


def projections_from_csv(source: pd.DataFrame, players: list[dict], name_col: str | None,
                         team_col: str | None, field_cols: dict[str, str | None]) -> tuple[dict, list[str]]:
    lookup = player_lookup_indices(players)
    result = {}
    unmatched = set()
    if not name_col:
        return result, ["Select the projection player-name column."]
    for _, row in source.iterrows():
        if pd.isna(row[name_col]) or not str(row[name_col]).strip():
            continue
        # Some optimizer projection exports append an index/footer row
        # (e.g. PLAYER="1", TEAM="4"). It is not an athlete.
        if not re.search(r"[a-zA-Z]", str(row[name_col])):
            continue
        key = identify_player(row[name_col], row[team_col] if team_col else None, lookup)
        if key is None:
            unmatched.add(str(row[name_col]).strip())
            continue
        record = {label: csv_numeric(row[column]) if column else float("nan")
                  for label, column in field_cols.items()}
        if key in result:
            # Conflicting projection rows should not silently select a value.
            for label, current in record.items():
                previous = result[key][label]
                if pd.notna(current) and pd.notna(previous) and not math.isclose(current, previous, abs_tol=1e-8):
                    result[key][label] = float("nan")
                    unmatched.add(f"Conflicting projection for {key[0]}: {label}")
                elif pd.isna(previous):
                    result[key][label] = current
        else:
            result[key] = record
    return result, sorted(unmatched)


def projection_eligible_players(source: pd.DataFrame, players: list[dict],
                                name_col: str | None, team_col: str | None,
                                allowed_codes: set[str]) -> tuple[list[dict], list[str]]:
    """The selected projection CSV is the ONLY roster for player views.

    Matching workbook players keep their nine stat groups. Players listed on
    projections but absent from the workbook get all-blank stat groups and may
    still receive their real optimizer lineup count/Roto score. Never borrow
    statistics from a player on a different team, or fabricate missing stats.
    """
    if not name_col:
        return [], ["Select the projection player-name column to build the eligible roster."]
    if not team_col:
        return [], ["Projection TEAM column is required for safe roster matching."]
    lookup = player_lookup_indices(players)
    player_index = {(p["name"], p["team"]): p for p in players}
    opponent_col = find_column(source.columns.tolist(),
                               ("OPP", "Opponent", "Opponent Team"))
    result = {}
    warnings = []
    for _, row in source.iterrows():
        raw = row[name_col]
        if pd.isna(raw) or not re.search(r"[a-zA-Z]", str(raw)):
            continue  # CSV footer, blank or numeric-only record
        code = normalized_team(row[team_col])
        if code not in allowed_codes:
            warnings.append(f"{str(raw).strip()}: projection team {code or '(missing)'} is not in the workbook slate")
            continue
        key = identify_player(raw, code, lookup)
        if key is not None:
            entry = dict(player_index[key])  # do not mutate the workbook entry
        else:
            display_name = re.sub(r"\s*\(\d{2,}\)\s*$", "", str(raw)).strip()
            entry = {"name": display_name, "team": code,
                     "groups": [[None] * 6 for _ in range(9)],
                     "_projection_only": True}
        if opponent_col:
            opponent_code = normalized_team(row[opponent_col])
            if opponent_code in allowed_codes and opponent_code != code:
                entry["_projection_opponent_code"] = opponent_code
        identity = (clean_player_name(entry["name"]), code)
        if identity not in result:
            result[identity] = entry
        elif (entry.get("_projection_opponent_code") and
              result[identity].get("_projection_opponent_code") and
              entry["_projection_opponent_code"] != result[identity]["_projection_opponent_code"]):
            # Conflicting CSV opponents are not used as a fallback.
            result[identity].pop("_projection_opponent_code", None)
            warnings.append(f"{entry['name']}: conflicting projection opponents")
    return list(result.values()), warnings


def suggest_projection_pool(projection_rows: pd.DataFrame, lineup_file,
                            slate_col: str, pools: list[str]) -> str:
    """Suggest a projection pool by roster-name overlap, not the CSV filename.

    The suggestion is purely a UI default. The user may select any pool.
    No projection player is inferred from a lineup that lacks the name.
    """
    if lineup_file is None or not pools:
        return "All pools"
    player_col = find_column(projection_rows.columns.tolist(), NAME_SYNONYMS)
    if player_col is None:
        return "All pools"
    try:
        lineup_rows = read_optional_csv(lineup_file.getvalue())
    except (ValueError, UnicodeError, OSError, pd.errors.EmptyDataError):
        return "All pools"
    slots = roster_position_columns(lineup_rows.columns.tolist())
    if slots:
        names = [value for col in slots for value in lineup_rows[col].dropna()]
    else:
        name_col = find_column(lineup_rows.columns.tolist(), NAME_SYNONYMS)
        if not name_col:
            return "All pools"
        names = list(lineup_rows[name_col].dropna())
    if not names:
        return "All pools"
    from collections import Counter
    def key(value):
        normalized = clean_player_name(value)
        return PLAYER_NAME_ALIASES.get(normalized, normalized)
    counts = Counter(key(name) for name in names)
    scores = {}
    for pool in pools:
        projected = projection_rows.loc[
            projection_rows[slate_col].fillna("").astype(str).str.strip().str.upper() == pool,
            player_col,
        ]
        scores[pool] = sum(counts.get(key(name), 0) for name in projected.dropna())
    best = max(scores, key=scores.get)
    # Never assume a pool if no overlap, or if two pools tie.
    if scores[best] <= 0 or list(scores.values()).count(scores[best]) > 1:
        return "All pools"
    return best


def source_csv_inputs(players: list[dict], team_names: list[str] | None = None,
                      lineup_file=None, projection_file=None) -> tuple[dict | None, dict, list[str], list[dict], bool]:
    """Read both CSVs automatically without the former CSV column-mapping panel.

    Required headers are detected from known aliases; report a clear warning
    rather than attaching one statistic to the wrong input. The selected
    projections pool is the exclusive eligibility source for player screens.
    """
    warnings = []
    eligible_players = []
    projections = {}
    counts = None
    projection_loaded = False
    codes = {TEAM_CODES[t] for t in (team_names or TEAM_CODES) if t in TEAM_CODES}

    if projection_file is not None:
        try:
            projection_csv = read_optional_csv(projection_file.getvalue())
            cols = projection_csv.columns.tolist()
            slate_col = find_column(cols, ("SLATE",))
            if slate_col:
                slate_values = projection_csv[slate_col].fillna("").astype(str).str.strip().str.upper()
                available = [v for v in ("MAIN", "LATE") if (slate_values == v).any()]
                if len(available) > 1:
                    recommended = suggest_projection_pool(projection_csv, lineup_file,
                                                          slate_col, available)
                    choices = ["All pools", *available]
                    pool_signature = hashlib.sha256(
                        projection_file.getvalue() +
                        (lineup_file.getvalue() if lineup_file is not None else b"")
                    ).hexdigest()[:14]
                    chosen = st.selectbox("Projection slate", choices,
                                          index=choices.index(recommended),
                                          key=f"csv_projection_slate_{pool_signature}",
                                          help="Choose MAIN or LATE if the projection CSV contains both pools.")
                    if chosen != "All pools":
                        projection_csv = projection_csv[slate_values == chosen]
            name_col = find_column(cols, NAME_SYNONYMS)
            team_col = find_column(cols, TEAM_SYNONYMS)
            field_cols = {label: find_column(cols, aliases)
                          for label, aliases in PROJECTION_SYNONYMS.items()}
            if not name_col or not team_col:
                st.error("Projection CSV needs a player-name and TEAM column. "
                         "The CSV mapping panel has been removed, so headers must be recognized automatically.")
            else:
                eligible_players, roster_warnings = projection_eligible_players(
                    projection_csv, players, name_col, team_col, codes)
                warnings.extend(f"Projection eligibility: {w}" for w in roster_warnings)
                projections, unmatched = projections_from_csv(
                    projection_csv, eligible_players, name_col, team_col, field_cols)
                warnings.extend(f"Projections: {w}" for w in unmatched)
                projection_loaded = bool(eligible_players)
                if projection_loaded:
                    st.caption(f"Projection roster: {len(eligible_players)} eligible players "
                               f"({sum(p.get('_projection_only', False) for p in eligible_players)} "
                               "without workbook NBA statistics).")
                else:
                    st.warning("No eligible players were matched from the projections CSV.")
            missing = [label for label, column in field_cols.items() if not column]
            if missing:
                st.warning("Projection columns not recognized: " + ", ".join(missing) +
                           ". Their Roto components remain blank.")
        except (ValueError, UnicodeError, OSError, pd.errors.EmptyDataError) as exc:
            st.error(f"Projection CSV error: {exc}")

    if lineup_file is not None:
        if not projection_loaded:
            st.warning("A valid projections CSV is required before optimized lineup counts "
                       "can be attached to eligible players.")
        else:
            try:
                lineup_csv = read_optional_csv(lineup_file.getvalue())
                cols = lineup_csv.columns.tolist()
                slots = roster_position_columns(cols)
                player_col = find_column(cols, NAME_SYNONYMS)
                count_col = find_column(cols, COUNT_SYNONYMS)
                if len(slots) >= 2:
                    mode = "Wide lineup rows"
                elif count_col and player_col:
                    mode = "Precomputed count"
                else:
                    mode = "Long player rows"
                if mode != "Wide lineup rows" and not player_col:
                    st.warning("Optimized lineups CSV has no recognized player-name or roster-slot columns. "
                               "Lineup appearances will be blank.")
                else:
                    counted, unmatched, occurrences = count_lineup_appearances(
                        lineup_csv, eligible_players, mode,
                        player_col, find_column(cols, TEAM_SYNONYMS),
                        find_column(cols, LINEUP_ID_SYNONYMS), count_col, slots)
                    if occurrences == 0 and len(lineup_csv):
                        st.warning("No optimized lineup entries match the selected projection pool. "
                                   "Lineup Points stay blank instead of becoming zero.")
                    else:
                        counts = counted
                        st.caption(f"Optimizer: {occurrences:,} eligible lineup appearances "
                                   f"across {len(counts)} players ({mode.lower()}).")
                    warnings.extend("Lineups: " + name for name in unmatched)
            except (ValueError, UnicodeError, OSError, pd.errors.EmptyDataError) as exc:
                st.error(f"Optimized lineups CSV error: {exc}")

    if warnings:
        with st.expander(f"Review {len(warnings)} unmatched players / CSV warnings"):
            st.write(warnings[:150])
            if len(warnings) > 150:
                st.caption("Showing the first 150 warnings.")
    return counts, projections, warnings, eligible_players, projection_loaded


def player_metadata(players: list[dict], counts: dict | None, projections: dict) -> pd.DataFrame:
    """Display raw projection inputs and all eight distinct Roto point components.

    A missing lineup CSV removes only Lineup Points. Other valid projection
    metrics still contribute; zero from a provided source is a genuine zero.
    """
    rows = []
    for entry in players:
        key = entry["name"], entry["team"]
        count = (counts.get(key, 0) if counts is not None else float("nan"))
        row = {"Player": key[0], "Team": key[1],
               "NBA Stats Source": ("Not in workbook" if entry.get("_projection_only")
                                    else "Workbook"),
               "Lineup Appearances": count}
        row.update(projections.get(key, {label: float("nan") for label in
                                         [*PROJECTION_LABELS, *PROP_PROJECTION_LABELS]}))
        row["Lineup Points"] = roto_points(count)
        for component, source in ROTO_COMPONENT_SOURCES.items():
            row[component] = roto_projection_points(source, row.get(source))
        available = [row[label] for label in ROTO_COMPONENT_LABELS if pd.notna(row[label])]
        row["Roto"] = float(sum(available)) if available else float("nan")
        row["Roto Components Present"] = len(available)
        row["Roto Status"] = (
            "Complete" if len(available) == len(ROTO_COMPONENT_LABELS) else
            "Partial" if available else "Unavailable"
        )
        rows.append(row)
    columns = ["Player", "Team", "NBA Stats Source", "Lineup Appearances",
               "Roto", *ROTO_COMPONENT_LABELS, "Roto Components Present",
               "Roto Status", *PROJECTION_LABELS, *PROP_PROJECTION_LABELS]
    return pd.DataFrame(rows, columns=columns)


def player_category_point_frames(long_data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return original stats, weighted per-stat points, and four scores / 20.

    Each player's comparison pool consists only of the projection-eligible
    players in the selected game scope. This is a relative player rating, not
    a fixed historical projection; opponent bonuses are added separately.

    Missing and non-finite values stay NaN. If all available values for a stat
    are the same, assign neutral points (not an arbitrary high rank). A stat
    absent from the entire workbook never becomes a fictional zero.
    """
    ids = ["Player", "Team", "Opponent"]
    selected = long_data[long_data["Stat"].isin(OVERALL_SCORE_STATS)].copy()
    if selected.duplicated(ids + ["Stat"]).any():
        raise ValueError("Duplicate player/opponent/stat rows in ranking input.")
    selected["Value"] = selected["Value"].map(valid_number_or_nan)
    wide = selected.pivot(index=ids, columns="Stat", values="Value")
    wide = wide.reindex(columns=list(OVERALL_SCORE_STATS))
    points = pd.DataFrame(index=wide.index)
    # An entirely missing column is unavailable to the entire comparison pool,
    # which is different from one player having a blank in an available column.
    universally_unavailable = {stat for stat in wide.columns if not wide[stat].notna().any()}
    for category, weights in PLAYER_SCORE_WEIGHTS.items():
        for stat, maximum in weights.items():
            if stat in universally_unavailable:
                points[stat] = pd.Series(float("nan"), index=wide.index)
                continue
            values = wide[stat]
            if values.notna().sum() == 0:
                points[stat] = pd.Series(float("nan"), index=wide.index)
            elif values.nunique(dropna=True) == 1:
                # Nobody is high or low if every player has an identical value.
                # Genuine zero production receives 0 points, never free neutral points.
                points[stat] = values.map(
                    lambda x: (float("nan") if pd.isna(x) else
                               0.0 if x <= 0 else float(maximum) / 2)
                )
            else:
                percentiles = values.rank(method="average", pct=True)
                points[stat] = percentiles.map(
                    lambda p: (float("nan") if pd.isna(p) else
                               0.0 if p <= 1 / 3 else
                               float(maximum) / 2 if p <= 2 / 3 else
                               float(maximum))
                )
                # A genuine reported 0 is not a missing value and never earns
                # positive performance points even if many players are tied at 0.
                points.loc[values <= 0, stat] = 0.0
    categories = pd.DataFrame(index=wide.index)
    for category, weights in PLAYER_SCORE_WEIGHTS.items():
        eligible = [stat for stat in weights if stat not in universally_unavailable]
        possible = sum(weights[stat] for stat in eligible)
        if not possible:
            categories[category] = pd.Series(float("nan"), index=wide.index)
            continue
        raw_points = points[eligible].sum(axis=1, min_count=1)
        # Scaling is only for entire stats missing across the selected slate.
        # Missing values for individual players never get free compensating points.
        categories[category] = (raw_points * 20.0 / possible).clip(lower=0, upper=20)
    return wide, points, categories


def build_player_overall_leaderboard(long_data: pd.DataFrame,
                                     metadata: pd.DataFrame,
                                     team_values: dict[str, dict] | None = None) -> pd.DataFrame:
    """Add raw-player tier scores, HALF the opposing team's category scores,
    and eight independent Roto components. Preserve partial/blank source data.

    The bonus is additive, never a multiplier. Base categories are /20;
    opponent bonuses are /10, so player categories are /30.
    """
    ids = ["Player", "Team", "Opponent"]
    bonus_cols = [f"{category} Matchup Bonus / 10" for category in PLAYER_SCORE_CATEGORIES]
    base_cols = [f"{category} Base / 20" for category in PLAYER_SCORE_CATEGORIES]
    cols = ["Rank", *ids, *PLAYER_SCORE_CATEGORIES,
            *base_cols, *bonus_cols, "Matchup Bonus / 40",
            "Matchup Categories Present", "Matchup Status",
            "Roto", "Overall Score", "Stats Present", "Score Status", "Missing Stats",
            "NBA Stats Source", "Lineup Appearances", *ROTO_COMPONENT_LABELS,
            "Roto Components Present", "Roto Status", *PROJECTION_LABELS]
    if long_data.empty:
        return pd.DataFrame(columns=cols)

    wide, _points, scores = player_category_point_frames(long_data)
    active_fields = list(wide.columns[wide.notna().any(axis=0)])
    scores["Stats Present"] = wide[active_fields].notna().sum(axis=1)
    scores["Missing Stats"] = wide[active_fields].apply(
        lambda row: ", ".join(row.index[row.isna()]), axis=1
    )
    result = scores.reset_index().merge(metadata, on=["Player", "Team"], how="left",
                                        validate="many_to_one")

    opposing_scores = None
    if team_values:
        # Shared with the Team Matchup Score tab: new 20-stat, four-tier team
        # categories power the additive player opponent bonuses.
        opposing_scores = build_overall_score_table(team_values).set_index("Team")

    for category, team_category in PLAYER_TO_TEAM_MATCHUP.items():
        base_col = f"{category} Base / 20"
        bonus_col = f"{category} Matchup Bonus / 10"
        result[base_col] = result[category]
        if opposing_scores is not None:
            opponent_score = pd.to_numeric(
                result["Opponent"].map(opposing_scores[team_category]), errors="coerce")
            result[bonus_col] = (opponent_score / 2.0).clip(lower=0, upper=10)
        else:
            result[bonus_col] = float("nan")
        # Missing base OR bonus does not erase the available component. Missing
        # values are not fabricated as 0; the category is flagged incomplete.
        result[category] = result[[base_col, bonus_col]].sum(axis=1, min_count=1).clip(0, 30)

    result["Matchup Bonus / 40"] = result[bonus_cols].sum(axis=1, min_count=1)
    result["Matchup Categories Present"] = result[bonus_cols].notna().sum(axis=1)
    result["Matchup Status"] = result["Matchup Categories Present"].map(
        lambda n: "Complete" if n == 4 else "Partial" if n else "Unavailable")

    columns_to_add = [*PLAYER_SCORE_CATEGORIES, "Roto"]
    result["Overall Score"] = result[columns_to_add].sum(axis=1, min_count=1)
    total_stats = len(active_fields)

    def score_status(row: pd.Series) -> str:
        if pd.isna(row["Overall Score"]):
            return "Unavailable"
        if row["Stats Present"] == 0:
            return ("Matchup / Roto only" if row["Matchup Categories Present"]
                    else "Roto only")
        if row["Stats Present"] != total_stats or row["Matchup Categories Present"] != 4:
            return "Partial"
        return "Complete"

    result["Score Status"] = result.apply(score_status, axis=1)
    result["Rank"] = result["Overall Score"].rank(ascending=False, method="min").astype("Int64")
    return result[cols].sort_values(["Overall Score", "Player", "Team"],
                                    ascending=[False, True, True], kind="stable",
                                    na_position="last").reset_index(drop=True)


def team_flag_summary(teams: dict[str, dict], selected: list[str]) -> pd.DataFrame:
    """Only team names and favorable/unfavorable opponent stat summaries."""
    rows = []
    for team in selected:
        if team not in teams:
            continue
        values = teams[team]
        positive = []
        negative = []
        for stat in TEAM_MATCHUP_THRESHOLDS:
            # 3PM is synthesized solely for Fantasy; keep the highlight view
            # focused on actual workbook source fields.
            if stat not in values:
                continue
            raw = valid_number_or_nan(values.get(stat))
            tier = matchup_tier(stat, raw)
            if tier == "Favorable":
                positive.append(f"{stat} {raw:.1f}")
            elif tier == "Unfavorable":
                negative.append(f"{stat} {raw:.1f}")
        rows.append({"Team": team,
                     "Favorable for opponent": ", ".join(positive) or "—",
                     "Unfavorable for opponent": ", ".join(negative) or "—"})
    return pd.DataFrame(rows, columns=["Team", "Favorable for opponent", "Unfavorable for opponent"])


def display_player_rankings(long_data: pd.DataFrame, metadata: pd.DataFrame,
                            team_values: dict[str, dict]) -> None:
    if long_data.empty:
        st.info("No player/opponent matches available. Check the selected daily slate.")
        return
    leaderboard = build_player_overall_leaderboard(long_data, metadata, team_values)
    st.caption("Four base player categories score 0–20 points apiece from eligible players' "
               "relative stat rankings. Each receives an ADDITIVE bonus equal to half "
               "the opposing defense's team matchup category score (up to 10 points). "
               "The final Paint, Shooting, Attacking and Fantasy scores are each /30. "
               "Attacking draws from Chaos, Fantasy from Production, Paint from Interior, "
               "and Shooting from Exterior. "
               "Roto has eight components, up to 80 points including salary, floor, and ceiling; "
               "maximum overall score = 200. No multipliers are applied.")
    st.caption("Matchup bonuses use the actual opposing team's four-category matchup scores. "
               "Daily schedule pairings take priority; the projections CSV's OPP column "
               "is used only as a fallback. If an opponent cannot be established, "
               "its matchup bonus stays blank; valid player and Roto points still count. "
               "FGM is derived from 2PM + 3PM when needed and may reflect rounding.")
    if long_data.loc[long_data["Stat"] == "C3PA", "Value"].map(valid_number_or_nan).notna().any():
        st.caption("Corner 3PA is included in Shooting: it earns 0/1/2 points "
                   "based on the eligible players' relative ranks.")
    else:
        st.caption("Corner 3PA is unavailable for the selected comparison pool. "
                   "Its two possible points are redistributed across the remaining "
                   "Shooting statistics so the category still has a 20-point ceiling.")
    tabs = st.tabs(["Overall Leaderboard", "Paint", "Shooting", "Attacking",
                    "Fantasy", "Roto", "Player Detail"])
    with tabs[0]:
        a, b, c = st.columns([1, 2, 1])
        with a:
            team_filter = st.selectbox("Team", ["All teams"] + sorted(leaderboard["Team"].unique()),
                                       key="leaderboard_team")
        with b:
            search = st.text_input("Search player", key="leaderboard_player_search")
        with c:
            partial = st.checkbox("Show partial / missing-data scores", value=True, key="leaderboard_partial")
        visible = leaderboard.copy()
        if team_filter != "All teams":
            visible = visible[visible["Team"] == team_filter]
        if search.strip():
            visible = visible[visible["Player"].str.contains(search.strip(), case=False,
                                                            regex=False, na=False)]
        if not partial:
            visible = visible[visible["Score Status"] == "Complete"]
        main_cols = [
            "Rank", "Player", "Team", "Opponent", "Player Price", "Overall Score",
            "Projected Minutes", "Lineup Appearances", "Price per Dollar",
            "Roto", "Paint", "Shooting", "Attacking", "Fantasy",
            "Matchup Bonus / 40", "Projected Fantasy Points", "Projected Floor",
            "Projected Ceiling", "Smash",
        ]
        # Friendly labels apply to the overall leaderboard only. Internal
        # scoring, exports and the category detail views stay unchanged.
        main_display = visible[main_cols].rename(columns={
            "Lineup Appearances": "Lineups", "Attacking": "Attack",
            "Matchup Bonus / 40": "Matchup Bonus",
            "Projected Fantasy Points": "Projected Fantasy",
            "Price per Dollar": "Price Per Dollar",
        }).copy()
        main_display["Opponent"] = main_display["Opponent"].map(
            lambda name: TEAM_CODES.get(name, name) if isinstance(name, str) else name)
        st.dataframe(main_display, hide_index=True, use_container_width=True,
                     height=min(740, 90 + 33 * max(len(main_display), 1)),
                     column_config={**one_decimal_config(main_display),
                                    "Rank": st.column_config.NumberColumn("Rank", format="%d"),
                                    "Lineups": st.column_config.NumberColumn("Lineups", format="%d")})
        complete = (leaderboard["Score Status"] == "Complete").sum()
        partial_count = (leaderboard["Score Status"] == "Partial").sum()
        without_nba = leaderboard["Score Status"].isin(["Roto only", "Matchup / Roto only"]).sum()
        st.caption(f"{len(leaderboard)} projection-eligible players: {complete} complete / "
                   f"{partial_count} partial / {without_nba} without NBA stats. "
                   "Scores include half the opposing team's matchup category points "
                   "where available. Missing fields stay blank; ranks are based on "
                   "the entire selected slate, not the on-screen filter.")
        with st.expander("Show base player points and opponent matchup bonuses"):
            breakdown = ["Player", "Team", "Opponent"]
            for category in PLAYER_SCORE_CATEGORIES:
                breakdown.extend([f"{category} Base / 20", f"{category} Matchup Bonus / 10", category])
            breakdown.extend(["Matchup Status", "Roto", "Overall Score"])
            st.dataframe(visible[breakdown], hide_index=True, use_container_width=True,
                         column_config=one_decimal_config(visible[breakdown]))
        st.download_button("Download player leaderboard CSV",
                           leaderboard.to_csv(index=False).encode("utf-8-sig"),
                           file_name="orange_nba_player_leaderboard.csv", mime="text/csv",
                           key="new_leaderboard_download")
    for category, tab in zip(PLAYER_SCORE_CATEGORIES, tabs[1:5]):
        with tab:
            stats = PLAYER_SCORE_CATEGORIES[category]
            selected = long_data[long_data["Stat"].isin(stats)]
            wide = (selected.pivot(index=["Player", "Team", "Opponent"], columns="Stat",
                                   values="Value").reindex(columns=list(stats)).reset_index())
            score_column = f"{category} / 30"
            base_col = f"{category} Base / 20"
            bonus_col = f"{category} Matchup Bonus / 10"
            scored = leaderboard[["Player", "Team", "Opponent", category,
                                  base_col, bonus_col]].rename(columns={category: score_column})
            wide = wide.merge(scored, on=["Player", "Team", "Opponent"], how="left",
                              validate="one_to_one")
            wide.insert(3, score_column, wide.pop(score_column))
            wide.insert(4, base_col, wide.pop(base_col))
            wide.insert(5, bonus_col, wide.pop(bonus_col))
            wide = wide.sort_values(score_column, ascending=False, na_position="last")
            st.caption("Score /30 = player stat score /20 + 50% of opposing team's "
                       "matching Matchup Score /20 (bonus /10). "
                       "Missing source fields stay blank, without preventing other points.")
            st.dataframe(wide, hide_index=True, use_container_width=True,
                         column_config=one_decimal_config(wide))
            with st.expander("Show each player's category points"):
                _raw, stat_points, _category_scores = player_category_point_frames(long_data)
                detail_points = stat_points[[field for field in stats if field in stat_points.columns]]
                detail_points = detail_points.reset_index()
                detail_points = detail_points.merge(scored, on=["Player", "Team", "Opponent"],
                                                    how="left", validate="one_to_one")
                detail_points = detail_points.sort_values(score_column, ascending=False,
                                                          na_position="last")
                st.dataframe(detail_points, hide_index=True, use_container_width=True,
                             column_config=one_decimal_config(detail_points))
    with tabs[5]:
        st.caption("Roto / 80 = Lineups / 20 + Smash / 10 "
                   "+ FPTS/$ / 10 + Salary / 10 + Minutes / 10 "
                   "+ Floor / 5 + Ceiling / 5 + Projected FPTS / 10. "
                   "Every source is scored independently. Missing source fields "
                   "stay blank; any available components still add to Roto.")
        roto_table = leaderboard[["Player", "Team", "Opponent", "Lineup Appearances",
                                  *ROTO_COMPONENT_LABELS, "Roto", "Roto Status",
                                  "Roto Components Present", *PROJECTION_LABELS]].sort_values(
                                      "Roto", ascending=False, na_position="last")
        st.dataframe(roto_table, hide_index=True, use_container_width=True,
                     column_config=one_decimal_config(roto_table))
        with st.expander("Roto point thresholds — all eight components"):
            st.markdown("**Optimized lineup appearances (maximum 20)**")
            st.dataframe(pd.DataFrame([
                {"Appearances": f"{low}+" if math.isinf(high) else f"{low}–{high}",
                 "Points": points}
                for low, high, points in ROTO_LINEUP_TIERS
            ]), hide_index=True, use_container_width=True)
            for source, tiers in ROTO_PROJECTION_TIERS.items():
                st.markdown(f"**{source} (maximum {tiers[-1][1]})**")
                st.dataframe(pd.DataFrame([
                    {"At least": cutoff, "Points": pts}
                    for cutoff, pts in tiers
                ]), hide_index=True, use_container_width=True)
            st.markdown("**Salary (maximum 10; lower salaries score more)**")
            st.dataframe(pd.DataFrame([
                {"Salary at most": cutoff, "Points": pts}
                for cutoff, pts in ROTO_SALARY_TIERS
            ] + [{"Salary at most": "Above $10,500", "Points": 0}]),
                         hide_index=True, use_container_width=True)
            st.caption("Smash scores 1 point starting at 9, up to 10 at 50+. "
                       "Minutes 39 to below 40 = 9; salary above $10,500 = 0.")
    with tabs[6]:
        players_df = leaderboard[["Player", "Team"]].drop_duplicates().sort_values(["Team", "Player"])
        if not players_df.empty:
            options = [(p, t) for p, t in players_df.itertuples(index=False, name=None)]
            chosen = st.selectbox("Inspect player", range(len(options)),
                                  format_func=lambda i: f"{options[i][0]} ({options[i][1]})",
                                  key="new_detail_player")
            name, team = options[chosen]
            row = leaderboard[(leaderboard["Player"] == name) & (leaderboard["Team"] == team)].iloc[0]
            detail = long_data[(long_data["Player"] == name) & (long_data["Team"] == team)]
            st.metric("Overall Score", f"{row['Overall Score']:.1f}" if pd.notna(row["Overall Score"]) else "—")
            st.caption(f"Rank #{int(row['Rank'])}" if pd.notna(row["Rank"]) else "Unranked")
            st.markdown("**Player score breakdown (base /20 + matchup bonus /10)**")
            player_breakdown = pd.DataFrame([
                {"Category": category,
                 "Player / 20": row[f"{category} Base / 20"],
                 "Opponent bonus / 10": row[f"{category} Matchup Bonus / 10"],
                 "Total / 30": row[category]}
                for category in PLAYER_SCORE_CATEGORIES
            ])
            st.dataframe(player_breakdown, hide_index=True, use_container_width=True,
                         column_config=one_decimal_config(player_breakdown))
            st.markdown("**Roto component scores**")
            roto_detail = pd.DataFrame([
                {"Component": label, "Points": row[label]}
                for label in ROTO_COMPONENT_LABELS
            ])
            st.dataframe(roto_detail, hide_index=True, use_container_width=True,
                         column_config=one_decimal_config(roto_detail))
            st.dataframe(detail[["Stat", "Value"]],
                         hide_index=True, use_container_width=True,
                         column_config={"Value": st.column_config.NumberColumn(format="%.1f")})
    st.download_button("Download original player scoring stats CSV",
                       long_data.to_csv(index=False).encode("utf-8-sig"),
                       file_name="orange_nba_player_scoring_stats.csv", mime="text/csv", key="new_long_download")

def uploaded_workbook_check(contents: bytes) -> str | None:
    """Validate Excel by its ZIP structure, never its filename/extension."""
    if not contents:
        return "The NBA workbook is empty."
    try:
        with zipfile.ZipFile(io.BytesIO(contents)) as archive:
            if "xl/workbook.xml" not in archive.namelist():
                return "This is not an Excel .xlsx workbook (missing Excel workbook data)."
    except (OSError, zipfile.BadZipFile):
        return "The NBA workbook must contain Excel .xlsx data. The filename does not matter."
    return None


def uploaded_csv_check(contents: bytes) -> str | None:
    """Validate CSV text by bytes and parsing, not extension/filename."""
    if not contents:
        return "The CSV file is empty."
    if contents[:4] in (b"PK\x03\x04", b"PK\x05\x06") or b"\x00" in contents[:4096]:
        return "This is not a text CSV export. Please select the original CSV data file."
    try:
        table = read_optional_csv(contents)
        if not len(table.columns) or table.empty:
            return "This CSV has no usable rows."
    except (UnicodeError, ValueError, OSError, pd.errors.EmptyDataError) as exc:
        return f"Unable to read the CSV text: {exc}"
    return None



# FINAL SIX BUILD PLAYER RANKINGS
# Five builds follow the corrected Numeric Thresholds scorecard exactly.
# Sixth Fantasy Score follows the user-supplied eight-stat scoring grid:
# three team/game inputs /4 + five player/CSV inputs /8 = 52.
# Neither the former Roto points nor unspecified CSV metrics enter rankings.
from decimal import Decimal, ROUND_HALF_UP

SIX_BUILD_TARGET_MAX = 6 * 52        # Six complete builds, 312 possible points

# Automatically transcribed from the scorecard's Numeric Thresholds tab.
# Each item: (Team|Player, source label, inclusive tier minimums).
SIX_BUILD_RULES = {'Interior Score': [('Team', 'OPP Paint Points', (43.1, 47.5, 52.0)),
                    ('Team', 'OPP Second Chance Points', (12.1, 15.5, 17.5)),
                    ('Team', 'OPP Offensive Rebounds', (10.1, 12.0, 14.0)),
                    ('Player', 'ORPG', (1.0, 1.5, 2.5, 3.5)),
                    ('Player', 'OREB%', (15.0, 20.5, 31.0, 49.0)),
                    ('Player', 'RA FGA', (3.0, 4.5, 6.0, 8.0)),
                    ('Player', '%PTS PITP', (48.0, 56.5, 64.5, 75.0)),
                    ('Player', '%PTS FT', (18.0, 21.0, 25.0, 30.0))],
 'Corner Score': [('Team', 'OPP FGM', (38.1, 41.0, 43.0)),
                  ('Team', 'OPP 3PTM', (12.1, 14.0, 15.5)),
                  ('Team', 'OPP C3PTA', (8.6, 10.0, 12.5)),
                  ('Player', '3PT FGM', (2.0, 2.5, 3.0, 4.0)),
                  ('Player', 'C3PTA', (1.5, 2.0, 2.5, 3.0)),
                  ('Player', '%FGA 3PT', (50.0, 56.0, 65.0, 71.0)),
                  ('Player', '%PTS 3PT', (47.0, 53.0, 59.0, 67.0)),
                  ('Player', 'FGM %AST', (70.0, 76.0, 84.0, 90.0))],
 'Attack Score': [('Team', 'OPP MR FGM + ATB 3PTM + FTM', (29.1, 32.0, 35.0)),
                  ('Team', 'OPP RA FGA', (20.6, 25.0, 29.5)),
                  ('Team', 'OPP NRA FGA', (13.6, 16.5, 20.5)),
                  ('Player', 'Paint FGM', (3.1, 4.0, 6.0, 7.5)),
                  ('Player', 'MR FGM + 3PTM + FTM', (4.5, 6.5, 8.5, 10.5)),
                  ('Player', 'RA FGA + NRA FGA', (5.5, 7.5, 9.5, 11.5)),
                  ('Player', 'MR FGA + 3PTA + FTA', (10.0, 13.0, 16.0, 20.0)),
                  ('Player', '2FGM %UAST', (54.0, 61.0, 68.0, 75.0))],
 'Chaos Score': [('Team', 'OPP FGA + Defensive Rebounds', (114.5, 119.5, 124.5)),
                 ('Team', 'OPP Fastbreak Points + Points Off Turnovers', (29.0, 33.0, 38.0)),
                 ('Team', 'OPP Steals + Blocks', (11.5, 13.5, 14.5)),
                 ('Player', 'STL%', (20.0, 24.0, 29.0, 35.0)),
                 ('Player', 'BLK%', (20.0, 28.0, 34.0, 45.0)),
                 ('Player', 'DREB%', (20.0, 27.0, 32.0, 40.0)),
                 ('Player', '%PTS Fastbreak', (14.0, 17.5, 23.0, 27.0)),
                 ('Player', '%PTS Off Turnovers', (16.0, 19.5, 24.5, 29.0))],
 'Star Score': [('Team', 'OPP Points', (108.0, 113.5, 120.0)),
                ('Team', 'OPP 2PTM + C3PTM', (29.0, 31.0, 33.0)),
                ('Team', 'OPP 3PTM + Assists', (37.0, 40.5, 44.0)),
                ('Player', 'USG%', (21.5, 25.0, 28.5, 34.0)),
                ('Player', 'ATB 3PTM + MR FGM + FTM', (4.5, 6.5, 9.0, 12.0)),
                ('Player', 'AST + MR FGA + 3PTA', (10.0, 12.5, 16.5, 20.0)),
                ('Player', 'AST%', (23.5, 26.5, 34.0, 42.5)),
                ('Player', 'FGM %UAST', (35.0, 44.5, 55.5, 67.0))],
 }

SIX_TEAM_SOURCES = {
    'OPP Paint Points': ('OPP PTS PAINT',),
    'OPP Second Chance Points': ('OPP PTS 2ND CHANCE',),
    'OPP Offensive Rebounds': ('OPP OREB',),
    'OPP FGM': ('OPP FGM',),
    'OPP 3PTM': ('OPP 3PM',),
    'OPP C3PTA': ('C3PA',),
    'OPP MR FGM + ATB 3PTM + FTM': ('MR FGM', 'ATB 3PM', 'OPP FTM'),
    'OPP RA FGA': ('RA FGA',),
    'OPP NRA FGA': ('NRA FGA',),
    'OPP FGA + Defensive Rebounds': ('OPP FGA', 'OPP DREB'),
    'OPP Fastbreak Points + Points Off Turnovers': ('OPP PTS FB', 'OPP PTS OFF TOV'),
    'OPP Steals + Blocks': ('OPP STL', 'OPP BLK'),
    'OPP Points': ('OPP PTS',),
    'OPP 2PTM + C3PTM': ('OPP 2PM', 'C3PM'),
    'OPP 3PTM + Assists': ('OPP 3PM', 'OPP AST'),
    'OPP PTS + DREB + OREB + AST': ('OPP PTS', 'OPP DREB', 'OPP OREB', 'OPP AST'),
}

SIX_PLAYER_SOURCES = {
    'ORPG': ('OREB',),
    'OREB%': ('%OREB',),
    'RA FGA': ('RA FGA',),
    '%PTS PITP': ('%PTS PITP',),
    '%PTS FT': ('%PTS FT',),
    '3PT FGM': ('3PM',),
    'C3PTA': ('C3PA',),
    '%FGA 3PT': ('%FGA 3PT',),
    '%PTS 3PT': ('%PTS 3PT',),
    'FGM %AST': ('FGM %AST',),
    'Paint FGM': ('RA FGM', 'NRA FGM'),
    'MR FGM + 3PTM + FTM': ('MR FGM', '3PM', 'FTM'),
    'RA FGA + NRA FGA': ('RA FGA', 'NRA FGA'),
    'MR FGA + 3PTA + FTA': ('MR FGA', '3PA', 'FTA'),
    '2FGM %UAST': ('2FGM %UAST',),
    'STL%': ('%STL',),
    'BLK%': ('%BLK',),
    'DREB%': ('%DREB',),
    '%PTS Fastbreak': ('%PTS FBPs',),
    '%PTS Off Turnovers': ('%PTS OffTO',),
    'USG%': ('USG%',),
    'ATB 3PTM + MR FGM + FTM': ('ATB 3PM', 'MR FGM', 'FTM'),
    'AST + MR FGA + 3PTA': ('AST', 'MR FGA', '3PA'),
    'AST%': ('%AST',),
    'FGM %UAST': ('FGM %UAST',),
}

# The sixth build is named Fantasy Score in the supplied replacement point grid.
# Five inputs use the optimizer / player projections CSVs, one compares projections
# with last-six-game workbook MPG, and three draw from team/game workbook fields.
SIX_BUILD_NAMES = [*SIX_BUILD_RULES, 'Fantasy Score']
SIX_FANTASY_RULES = [
    ('Team', 'Game Total', 4),
    ('Team', 'Spread', 4),
    ('Team', 'OPP PTS + DREB + OREB + AST', 4),
    ('Player / CSV', 'Projected Minutes - MPG', 8),
    ('Player / CSV', '# of Lineups', 8),
    ('Player / CSV', 'Player Salary', 8),
    ('Player / CSV', 'Price Per Dollar', 8),
    ('Player / CSV', 'Projected Fantasy Points', 8),
]
assert len(SIX_BUILD_RULES) == 5 and len(SIX_FANTASY_RULES) == 8
assert sum(maximum for _, _, maximum in SIX_FANTASY_RULES) == 52
SIX_FANTASY_RAW_LABELS = (
    'Game Total', 'Spread', 'OPP PTS + DREB + OREB + AST',
    'Last 6 MPG', 'Projected Minutes', 'Projected Minutes - MPG',
    '# of Lineups', 'Player Salary', 'Price Per Dollar',
    'Projected Fantasy Points',
)


def six_fantasy_score_points(label: str, value: object) -> float:
    """Exact inclusive new user tiers, with decimal comparisons after rounding.

    Spread is the ABSOLUTE margin: +3.5 and -3.5 earn the same points.
    Thresholds described at one decimal are scored on half-up rounded values.
    For the two ambiguous one-decimal gaps, 9.6 spread earns 0 and a
    $9,901-$9,999 salary earns 0; this avoids awarding an unstated higher tier.
    """
    n = six_round_1(value)
    if pd.isna(n) or n < 0 and label not in ('Spread', 'Projected Minutes - MPG'):
        return float('nan')
    if label == 'Game Total':
        return float(4 if n >= 234.5 else 2 if n >= 225.5 else 1 if n >= 216.5 else 0)
    if label == 'Spread':
        n = abs(n)
        return float(4 if n <= 2.5 else 2 if n <= 5.5 else 1 if n <= 9.5 else 0)
    if label == 'OPP PTS + DREB + OREB + AST':
        return float(4 if n >= 192.5 else 2 if n >= 183.5 else 1 if n > 174.5 else 0)
    if label == 'Projected Minutes - MPG':
        return float(8 if n >= 16.0 else 4 if n >= 12.0 else 2 if n >= 8.5 else 1 if n >= 5.5 else 0)
    if label == '# of Lineups':
        if n != int(n):
            return float('nan')
        return float(8 if n >= 150 else 4 if n >= 85 else 2 if n >= 35 else 1 if n >= 1 else 0)
    if label == 'Player Salary':
        return float(8 if n <= 3600 else 4 if n <= 5400 else 2 if n <= 7500 else 1 if n <= 9900 else 0)
    if label == 'Price Per Dollar':
        return float(8 if n >= 7.0 else 4 if n >= 5.0 else 2 if n >= 4.0 else 1 if n >= 3.1 else 0)
    if label == 'Projected Fantasy Points':
        return float(8 if n >= 65.0 else 4 if n >= 51.0 else 2 if n >= 37.0 else 1 if n >= 25.0 else 0)
    raise ValueError('Unknown Fantasy Score statistic: ' + label)


SIX_FANTASY_THRESHOLD_ROWS = [
    ('Team', 'Game Total', '<216.5', '216.5–225.4', '225.5–234.4', '≥234.5', '—'),
    ('Team', 'Spread (absolute)', '≥9.6', '5.6–9.5', '2.6–5.5', '≤2.5', '—'),
    ('Team', 'OPP PTS + DREB + OREB + AST', '≤174.5', '174.6–183.4', '183.5–192.4', '≥192.5', '—'),
    ('Player / CSV', 'Projected Minutes - MPG', '<5.5', '5.5–8.4', '8.5–11.9', '12.0–15.9', '≥16.0'),
    ('Player / CSV', '# of Lineups', '0', '1–34', '35–84', '85–149', '≥150'),
    ('Player / CSV', 'Player Salary', '>$9,900', '$7,501–9,900', '$5,401–7,500', '$3,601–5,400', '≤$3,600'),
    ('Player / CSV', 'Price Per Dollar', '≤3.0', '3.1–3.9', '4.0–4.9', '5.0–6.9', '≥7.0'),
    ('Player / CSV', 'Projected Fantasy Points', '<25.0', '25.0–36.9', '37.0–50.9', '51.0–64.9', '≥65.0'),
]


def six_stat_number(values: dict, sources: tuple[str, ...], *, team: bool) -> float:
    numbers = [team_stat_value(values, source) if team else valid_number_or_nan(values.get(source))
               for source in sources]
    if any(pd.isna(number) for number in numbers):
        return float('nan')
    return float(sum(numbers))


def six_round_1(value: object) -> float:
    """Excel-compatible ROUND(value,1): halves away from zero."""
    number = valid_number_or_nan(value)
    if pd.isna(number):
        return float('nan')
    return float(Decimal(str(number)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP))


def six_score_value(value: object, inclusive_minimums: tuple[float, ...]) -> float:
    number = six_round_1(value)
    if pd.isna(number):
        return float('nan')
    tiers = (1, 2, 4, 8) if len(inclusive_minimums) == 4 else (1, 2, 4)
    return float(next((pts for cutoff, pts in reversed(list(zip(inclusive_minimums, tiers)))
                       if number >= cutoff), 0))


def six_build_calculate(eligible_players: list[dict], team_values: dict[str, dict],
                        player_boxes: list[dict], metadata: pd.DataFrame,
                        opponent_pairs: dict[str, str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Six /52 scores plus a complete 48-input calculation trace.

    Unavailable scores remain missing. Rank totals sum only valid components,
    with an explicit completion indicator. Game total and spread describe the
    PLAYER'S game/team; defense-allowed sums come from the OPPOSING team.
    """
    team_by_code = {TEAM_CODES[name]: name for name in team_values if name in TEAM_CODES}
    meta_by_player = metadata.set_index(['Player', 'Team']) if not metadata.empty else None
    overall_rows, trace_rows = [], []
    for entry in eligible_players:
        name, code = entry['name'], entry['team']
        own_team = team_by_code.get(code, '')
        own_team_values = team_values.get(own_team, {})
        opponent = opponent_pairs.get(own_team, '') if own_team else ''
        if not opponent:
            candidate = team_by_code.get(entry.get('_projection_opponent_code', ''), '')
            if candidate and candidate != own_team:
                opponent = candidate
        opposing_defense = team_values.get(opponent, {})
        values, conflicts = player_raw_stat_values(entry, player_boxes)
        metadata_key = (name, code)
        meta = (meta_by_player.loc[metadata_key] if meta_by_player is not None
                and metadata_key in meta_by_player.index else None)
        row = {'Player': name, 'Team': code, 'Opponent': TEAM_CODES.get(opponent, ''),
               'NBA Stats Source': ('Not in workbook' if entry.get('_projection_only') else 'Workbook')}
        completeness = []
        for build, rules in SIX_BUILD_RULES.items():
            points_present = []
            for kind, label, minimums in rules:
                sources = SIX_TEAM_SOURCES[label] if kind == 'Team' else SIX_PLAYER_SOURCES[label]
                value = six_stat_number(opposing_defense if kind == 'Team' else values, sources,
                                        team=kind == 'Team')
                points = six_score_value(value, minimums)
                if pd.notna(points):
                    points_present.append(points)
                trace_rows.append({
                    'Player': name, 'Team': code, 'Opponent': row['Opponent'], 'Build': build,
                    'Type': kind, 'Statistic': label,
                    'Actual Value': six_round_1(value), 'Points': points,
                    'Max Points': 4 if kind == 'Team' else 8,
                    'Source': 'Opponent team workbook' if kind == 'Team' else 'Player workbook',
                })
            row[build] = float(sum(points_present)) if points_present else float('nan')
            row[build + ' Inputs'] = len(points_present)
            completeness.append(len(points_present))

        def raw_meta(field):
            return valid_number_or_nan(meta.get(field)) if meta is not None else float('nan')

        last_mpg = valid_number_or_nan(values.get('MPG'))
        projected_mins = raw_meta('Projected Minutes')
        minutes_delta = (projected_mins - last_mpg if pd.notna(last_mpg)
                         and pd.notna(projected_mins) else float('nan'))
        fantasy_values = {
            'Game Total': valid_number_or_nan(own_team_values.get('Game Total')),
            'Spread': valid_number_or_nan(own_team_values.get('Spread')),
            'OPP PTS + DREB + OREB + AST': six_stat_number(
                opposing_defense, SIX_TEAM_SOURCES['OPP PTS + DREB + OREB + AST'], team=True),
            'Projected Minutes - MPG': minutes_delta,
            '# of Lineups': raw_meta('Lineup Appearances'),
            'Player Salary': raw_meta('Player Price'),
            'Price Per Dollar': raw_meta('Price per Dollar'),
            'Projected Fantasy Points': raw_meta('Projected Fantasy Points'),
        }
        row.update(fantasy_values)
        row['Last 6 MPG'] = last_mpg
        row['Projected Minutes'] = projected_mins
        fantasy_points = []
        for kind, label, maximum in SIX_FANTASY_RULES:
            value = fantasy_values[label]
            points = six_fantasy_score_points(label, value)
            if pd.notna(points):
                fantasy_points.append(points)
            trace_rows.append({
                'Player': name, 'Team': code, 'Opponent': row['Opponent'], 'Build': 'Fantasy Score',
                'Type': kind, 'Statistic': label, 'Actual Value': six_round_1(value),
                'Points': points, 'Max Points': maximum,
                'Source': ('Own-team ESPN workbook' if label in ('Game Total', 'Spread') else
                           'Opponent team workbook' if kind == 'Team' else
                           'Player workbook + projections CSV' if label == 'Projected Minutes - MPG' else
                           'Optimizer CSV' if label == '# of Lineups' else 'Player projections CSV'),
            })
        row['Fantasy Score'] = (float(sum(fantasy_points)) if fantasy_points
                                else float('nan'))
        row['Fantasy Score Inputs'] = len(fantasy_points)
        completeness.append(len(fantasy_points))
        scored = [row[b] for b in SIX_BUILD_NAMES if pd.notna(row[b])]
        row['Overall Score / 312'] = float(sum(scored)) if scored else float('nan')
        row['Stats Present'] = int(sum(completeness))
        row['Score Status'] = ('Complete' if sum(completeness) == 48 else
                               'Partial' if sum(completeness) else 'Unavailable')
        if conflicts:
            row['Conflicting Fields'] = ', '.join(sorted(set(conflicts)))
        overall_rows.append(row)
    leaderboard = pd.DataFrame(overall_rows)
    traces = pd.DataFrame(trace_rows)
    if not leaderboard.empty:
        leaderboard = leaderboard.sort_values(['Overall Score / 312', 'Player', 'Team'],
                                              ascending=[False, True, True],
                                              kind='stable', na_position='last').reset_index(drop=True)
        leaderboard.insert(0, 'Rank', leaderboard['Overall Score / 312'].rank(
            method='min', ascending=False).astype('Int64'))
    return leaderboard, traces


def six_build_threshold_table(build: str) -> pd.DataFrame:
    if build == 'Fantasy Score':
        return pd.DataFrame([
            {'Type': kind, 'Statistic': label, '0 points': zero, '1 point': one,
             '2 points': two, '4 points': four, '8 points': eight}
            for kind, label, zero, one, two, four, eight in SIX_FANTASY_THRESHOLD_ROWS
        ])
    rows = []
    for kind, label, cutoffs in SIX_BUILD_RULES[build]:
        row = {'Type': kind, 'Statistic': label, '0 points': f'< {cutoffs[0]:g}',
               '1 point': f'≥ {cutoffs[0]:g}', '2 points': f'≥ {cutoffs[1]:g}',
               '4 points': f'≥ {cutoffs[2]:g}',
               '8 points': f'≥ {cutoffs[3]:g}' if len(cutoffs) == 4 else '—'}
        rows.append(row)
    return pd.DataFrame(rows)


def display_six_build_rankings(eligible_players: list[dict], team_values: dict[str, dict],
                               player_boxes: list[dict], metadata: pd.DataFrame,
                               opponents: dict[str, str]) -> None:
    if not eligible_players:
        st.info('Upload a valid projections CSV to populate the eligible player rankings.')
        return
    board, trace = six_build_calculate(eligible_players, team_values, player_boxes,
                                       metadata, opponents)
    if board.empty:
        st.info('No eligible players found in the chosen game scope.')
        return
    st.caption('Final Six Builds: Interior, Corner, Attack, Chaos, Star and Fantasy are '
               'each worth up to 52 points. Maximum overall score: **312**. '
               'Fantasy uses Game Total, absolute Spread and the opponent aggregate '
               '(/4 each), plus five CSV/player inputs (/8 each). '
               'Scores use one-decimal rounded source values.')
    names = SIX_BUILD_NAMES
    pages = st.tabs(['Overall Leaderboard', *names])
    with pages[0]:
        a, b = st.columns([1, 2])
        with a:
            choose_team = st.selectbox('Team', ['All teams', *sorted(board['Team'].dropna().unique())],
                                       key='six_leaderboard_team')
        with b:
            search = st.text_input('Search player', key='six_leaderboard_search')
        visible = board
        if choose_team != 'All teams':
            visible = visible[visible['Team'] == choose_team]
        if search.strip():
            visible = visible[visible['Player'].str.contains(search.strip(), regex=False,
                                                              case=False, na=False)]
        cols = ['Rank', 'Player', 'Team', 'Opponent', 'Overall Score / 312',
                *names, 'Stats Present', 'Score Status']
        st.dataframe(visible[cols], hide_index=True, use_container_width=True,
                     height=min(760, 85 + 33 * max(len(visible), 1)),
                     column_config=one_decimal_config(visible[cols]))
        st.caption('Overall is calculated from 48 scoring inputs (8 per build). '
                   'Missing values remain unscored and are flagged Partial, not treated '
                   'as actual zeros. Ranks are calculated on the full selected game scope.')
        st.download_button('Download Six-Build Leaderboard CSV',
                           board.to_csv(index=False).encode('utf-8-sig'),
                           file_name='orange_nba_final_six_build_leaderboard.csv',
                           mime='text/csv', key='six_leaderboard_download')
    for page, build in zip(pages[1:], names):
        with page:
            selected = trace[trace['Build'] == build]
            if build != 'Fantasy Score':
                st.caption('Three opponent team stats award 0/1/2/4 points each; '
                           'five player stats award 0/1/2/4/8 points each. Maximum: 52.')
            else:
                st.caption('Game Total and absolute Spread use the player’s own game; '
                           'the combined opponent statistic uses the opposing defense. '
                           'Projected Minutes - MPG = CSV projected minutes minus last-six '
                           'workbook MPG. Three 4-point stats and five 8-point stats = 52.')
            details = board[['Player', 'Team', 'Opponent', build]].rename(
                columns={build: 'Score / 52'})
            # Pivot even if every score is missing for a particular stat.
            points_wide = selected.pivot_table(index=['Player', 'Team', 'Opponent'],
                                               columns='Statistic', values='Points', aggfunc='first')
            labels = ([label for _kind, label, _tiers in SIX_BUILD_RULES[build]]
                      if build != 'Fantasy Score' else
                      [label for _kind, label, _maximum in SIX_FANTASY_RULES])
            for label in labels:
                if label not in points_wide:
                    points_wide[label] = float('nan')
            points_wide = points_wide[labels].reset_index()
            details = details.merge(points_wide, on=['Player', 'Team', 'Opponent'], how='left')
            if build == 'Fantasy Score':
                details = details.merge(board[['Player', 'Team', 'Last 6 MPG', 'Projected Minutes']],
                                        on=['Player', 'Team'], how='left')
            details = details.sort_values('Score / 52', ascending=False,
                                          kind='stable', na_position='last')
            st.dataframe(details, hide_index=True, use_container_width=True,
                         column_config=one_decimal_config(details))
            with st.expander('Show actual values and awarded points for each statistic'):
                st.dataframe(selected, hide_index=True, use_container_width=True,
                             column_config=one_decimal_config(selected))
            with st.expander('Scorecard point thresholds'):
                st.dataframe(six_build_threshold_table(build), hide_index=True,
                             use_container_width=True)
            st.download_button('Download ' + build + ' details CSV',
                               selected.to_csv(index=False).encode('utf-8-sig'),
                               file_name='orange_nba_' + csv_col_key(build) + '_details.csv',
                               mime='text/csv', key='six_details_' + csv_col_key(build))


# PAGE LAYOUT
st.set_page_config(page_title="Orange NBA Viewer", page_icon="🏀", layout="wide",
                   initial_sidebar_state="collapsed")
st.markdown("""
<style>
.block-container {max-width:1500px; padding-top:1.3rem; padding-bottom:3rem;}
.nba-kicker {color:#f97316;font-size:.8rem;font-weight:750;letter-spacing:.12em;text-transform:uppercase;}
.stTabs [data-baseweb="tab-list"] {gap:.3rem;}
.stTabs [data-baseweb="tab"] {font-weight:650;white-space:nowrap;}
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="nba-kicker">Orange NBA Viewer</div>', unsafe_allow_html=True)
st.title("NBA Stats Explorer")
st.caption("Daily game-slate lookup · Team matchup /80 · Final Six Player Builds (six /52; overall /312).")

st.subheader("Upload daily NBA files")
st.caption("Files are accepted based on their actual content, **not their filenames**. "
           "Upload an Excel .xlsx workbook and two CSV files. All three selectors "
           "are always visible; you can upload in any order.")
workbook_col, optimizer_col, projections_col = st.columns(3)
with workbook_col:
    uploaded = st.file_uploader(
        "1. NBA stats workbook", type=None, key="nba_workbook",
        help="Select an Excel .xlsx file with Player Stats and Team Stats sheets. "
             "Any filename is accepted.")
with optimizer_col:
    lineup_upload = st.file_uploader(
        "2. Optimized lineups CSV", type=None, key="optimized_csv",
        help="Select your optimized lineups CSV. Any filename is accepted.")
with projections_col:
    projection_upload = st.file_uploader(
        "3. Player projections CSV", type=None, key="projection_csv",
        help="Select your projections CSV. Any filename is accepted.")

# Streamlit's file picker uses no extension/name filter at all. The content
# checks below make sure a file is readable in its intended upload slot.
for source_label, selected_file, check in (
    ("NBA workbook", uploaded, uploaded_workbook_check),
    ("Optimized lineups", lineup_upload, uploaded_csv_check),
    ("Roto projections", projection_upload, uploaded_csv_check),
):
    if selected_file is not None:
        issue = check(selected_file.getvalue())
        if issue:
            st.error(f"{source_label}: {issue}")
        else:
            st.success(f"{source_label} accepted ({len(selected_file.getvalue()):,} bytes).")

# Invalid CSVs are not used for calculations, but they do not hide the other
# uploaders, and valid files can still be replaced without restarting the app.
if lineup_upload is not None and uploaded_csv_check(lineup_upload.getvalue()):
    lineup_upload = None
if projection_upload is not None and uploaded_csv_check(projection_upload.getvalue()):
    projection_upload = None
if uploaded is None:
    st.info("Select your NBA .xlsx workbook to calculate scores. Your CSVs can "
            "already be uploaded, regardless of their filenames.")
    st.stop()
if uploaded_workbook_check(uploaded.getvalue()):
    st.info("Please select a valid Excel .xlsx workbook in the first upload box.")
    st.stop()
try:
    data = cached_workbook(uploaded.getvalue())
    team_values = team_values_for_scoring(data["Team Stats"])
except (OSError, ValueError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
    st.error(f"Unable to read this NBA workbook: {exc}")
    st.caption("Expected Team Stats: 25 scoring columns, plus optional columns Z–AK "
               "for matchup projections and ESPN odds. Player Stats: 56 original "
               "columns plus optional BE=FTM and BF=%PF. Grouped headers are "
               "in row 1 and column headers in row 2.")
    st.stop()
team_names = data["teamNames"]
players = data["Player Stats"]["entries"]
st.success(f"Workbook loaded: {len(team_names)} teams and {len(players)} players.")

selected_slate_pairs = matchup_selection_ui(team_names)
slate_pair_labels = {game_label(a, b): (a, b) for a, b in selected_slate_pairs}
slate_team_names = [team for pair in selected_slate_pairs for team in pair]
# CSV detection is automatic; there is no expandable column-mapping form.
counts, projections, csv_warnings, eligible_players, has_projection_pool = source_csv_inputs(
    players, team_names, lineup_file=lineup_upload, projection_file=projection_upload)
metadata = player_metadata(eligible_players, counts, projections)
visible_player_data = {**data["Player Stats"], "entries": eligible_players}

team_main, score_main, player_main, player_adjust_main = st.tabs(
    ["Team Stats", "Matchup Score", "Player Stats", "Player Rankings"]
)

with team_main:
    st.subheader("Team Stats — Box Zero + Boxes 1–4")
    st.caption("Positive / favorable means that defense allows a better opportunity for its "
               "OPPONENT. Negative / unfavorable means the defense allows less. "
               "These flags use your existing tiers, not yet-to-be-defined statistical outliers.")
    team_options = ["All teams", "Teams on selected slate", *slate_pair_labels, "Choose specific teams"]
    team_filter = st.selectbox("Show teams / game", team_options, key="team_filter")
    if team_filter == "Choose specific teams":
        selected_teams = st.multiselect("Teams to display", team_names,
                                        default=team_names, key="selected_team_names")
    elif team_filter == "Teams on selected slate":
        selected_teams = slate_team_names
    elif team_filter == "All teams":
        selected_teams = team_names
    else:
        selected_teams = list(slate_pair_labels[team_filter])

    st.markdown("### Team highlights — opponent opportunities")
    highlights = team_flag_summary(team_values, selected_teams)
    if highlights.empty:
        st.info("No teams selected.")
    else:
        st.dataframe(highlights, hide_index=True, use_container_width=True,
                     height=min(620, 85 + 36 * len(highlights)),
                     column_config={"Favorable for opponent": st.column_config.TextColumn(width="large"),
                                    "Unfavorable for opponent": st.column_config.TextColumn(width="large")})
        st.caption("Only the selected team names and their favorable/unfavorable "
                   "opponent statistics are shown; neutral statistics are omitted.")
    st.divider()
    st.markdown("### Team Box Zero — opponent totals and shot zones")
    show_team_box_zero(team_values, selected_teams)
    for index, box in enumerate(data["Team Stats"]["boxes"]):
        st.divider()
        st.markdown(f"### Team Box {index + 1} — six stats")
        frame = six_stat_table(data["Team Stats"], "Team Stats", index)
        frame = frame[frame["Team"].isin(selected_teams)]
        show_six_stat_tab(frame, section="teams", box_name=box["name"])
    odds_table = espn_game_odds_table(data["Team Stats"], selected_teams)
    if not odds_table.empty:
        st.divider()
        st.markdown("### ESPN game odds and expected totals")
        st.caption("These are the lines and market-implied expected totals saved "
                   "in the uploaded Excel file, not actual scores or live odds.")
        st.dataframe(odds_table, hide_index=True, use_container_width=True,
                     column_config=one_decimal_config(odds_table))

with score_main:
    st.subheader("Matchup Score")
    st.caption("Four scoring categories (Interior, Exterior, Production, Chaos), "
               "each out of 20. Five defense-allowed stats in each category earn "
               "0 / 1 / 2 / 4 points from the exact four-tier thresholds; "
               "20 unique statistics make up the overall /80 matchup score.")
    score_options = ["All teams", "Teams on selected slate", *slate_pair_labels, "Choose specific teams"]
    score_filter = st.selectbox("Show teams / matchup", score_options, key="score_team_filter")
    if score_filter == "Choose specific teams":
        scoring_teams = st.multiselect("Teams to score", team_names, default=team_names,
                                       key="scoring_team_names")
    elif score_filter == "All teams":
        scoring_teams = team_names
    elif score_filter == "Teams on selected slate":
        scoring_teams = slate_team_names
    else:
        scoring_teams = list(slate_pair_labels[score_filter])
    try:
        st.markdown("### Alpha — 12 actual statistics + Total / 50")
        st.caption("Alpha /50 keeps its existing three-tier thresholds; the four-tier "
                   "0 / 1 / 2 / 4 scoring applies to the four matchup categories below.")
        show_alpha_table(team_values, scoring_teams)
        st.divider()
        all_scores = build_overall_score_table(team_values)
        visible_scores = all_scores[all_scores["Team"].isin(scoring_teams)]
        st.markdown("### Overall matchup score / 80")
        visible_scores = visible_scores.sort_values("Overall / 80", ascending=False,
                                                    na_position="last", kind="stable")
        # Coverage and completeness stay in the calculations, but not the table.
        overall_display = visible_scores.drop(columns=["Stats Present", "Status", "Missing"],
                                              errors="ignore")
        st.dataframe(overall_display, hide_index=True, use_container_width=True,
                     column_config=one_decimal_config(overall_display))
        st.caption("Higher totals mean the defense presents a more favorable matchup "
                   "for its opponent. Missing stats are not silently treated as zeros.")
        score_tabs = st.tabs(list(MATCHUP_SCORE_GROUPS))
        for category, tab in zip(MATCHUP_SCORE_GROUPS, score_tabs):
            with tab:
                st.markdown(f"**{category} — five source stats, 20 points maximum**")
                category_df = build_matchup_score_table(team_values, category)
                category_df = category_df[category_df["Team"].isin(scoring_teams)]
                show_score_table(category_df, category)
                with st.expander("Actual defense-allowed values behind these points"):
                    source_df = pd.DataFrame([
                        {"Team": t, **{label: team_stat_value(team_values[t], field)
                                       for label, field, _ in MATCHUP_SCORE_GROUPS[category]}}
                        for t in scoring_teams if t in team_values
                    ])
                    if not source_df.empty:
                        styled = source_df.style
                        for label, source, _ in MATCHUP_SCORE_GROUPS[category]:
                            styled = styled.map(lambda value, field=source: matchup_score_tier_style(field, value),
                                                subset=[label])
                        st.dataframe(styled.format({c: "{:.1f}" for c in source_df.columns[1:]},
                                                  na_rep="—"), hide_index=True, use_container_width=True)
    except ValueError as exc:
        st.error(f"Matchup Score unavailable: {exc}")

with player_main:
    st.subheader("Player Stats — Interior, Exterior, Tendency, Production, Fantasy, DFS & Props")
    st.caption("Historical NBA averages come from your workbook. DFS and player prop projections "
               "come from the selected projections CSV; # of Lineups is the raw appearance count "
               "in the optimized lineups CSV. On new workbooks, Expected Team/Game Totals "
               "come from the uploaded ESPN odds / implied-total columns, not final scores. "
               "Older workbooks without odds can use their Expected Offense estimates.")
    if not has_projection_pool:
        st.info("Upload the Roto projections CSV above to show eligible players. "
                "Players absent from the selected projection pool are excluded.")
    else:
        st.caption(f"Showing {len(eligible_players)} projection-eligible players; "
                   f"{sum(p.get('_projection_only', False) for p in eligible_players)} "
                   "have no NBA workbook stat row.")
    a, b = st.columns([1, 2])
    with a:
        chosen_team = st.selectbox(
            "Team filter", ["All teams"] + sorted({p["team"] for p in eligible_players}),
            key="player_filter_team")
    with b:
        search = st.text_input("Find player", placeholder="Search name", key="player_search")
    player_tabs = st.tabs(list(PLAYER_CHART_GROUPS))
    for chart_name, tab in zip(PLAYER_CHART_GROUPS, player_tabs):
        with tab:
            chart, missing = player_chart_table(
                visible_player_data, chart_name, metadata, data["Team Stats"], selected_slate_pairs)
            if chosen_team != "All teams":
                chart = chart[chart["Team"] == chosen_team]
            if search.strip():
                chart = chart[chart["Player"].str.contains(
                    search.strip(), case=False, regex=False, na=False)]
            if chart_name == "DFS":
                st.caption("# of Lineups is a raw optimizer count; it is not the 0–20 Roto lineup score. "
                           "If no valid matching optimized lineups CSV was supplied, the count stays blank.")
            if chart_name == "Props":
                st.caption("Expected Game Total = ESPN Game Total (AH). "
                           "Expected Team Total = ESPN-implied team points (AJ). "
                           "Both are pregame projections from your uploaded workbook, "
                           "not actual results. Legacy workbooks can fall back to "
                           "Expected Offense estimates.")
            if chart_name == "Fantasy":
                st.caption("PF% reads the Player Stats workbook's %PF field (BF), "
                           "separately from personal fouls per game (PF).")
            show_player_chart(chart, missing, chart_name)
            if chart_name == "Props":
                visible_codes = set(chart["Team"].dropna())
                visible_teams = [name for name in team_names
                                 if TEAM_CODES.get(name) in visible_codes]
                odds_table = espn_game_odds_table(data["Team Stats"], visible_teams)
                if not odds_table.empty:
                    with st.expander("See ESPN spreads, game times, venues and expected opponent totals"):
                        st.dataframe(odds_table, hide_index=True, use_container_width=True,
                                     column_config=one_decimal_config(odds_table))

with player_adjust_main:
    st.subheader("Player Rankings — Final Six Builds")
    st.caption("Only players in the selected projection pool appear. Six fixed-threshold "
               "builds each score up to 52 points (overall /312). Fantasy Score includes "
               "the specified sportsbook game total / spread, opponent aggregate, "
               "and five player / optimizer CSV metrics. The former Roto /80 and "
               "player percentile scoring do not enter this leaderboard.")
    game_options = ["Entire selected slate", *slate_pair_labels]
    selected_game = st.selectbox("Game scope", game_options,
                                 key="player_adjustment_selection_mode")
    pairs = (selected_slate_pairs if selected_game == "Entire selected slate"
             else [slate_pair_labels[selected_game]])
    opponents = {}
    for away, home in pairs:
        opponents[away] = home
        opponents[home] = away
    if not opponents:
        st.info("No selected schedule matchups; projection OPP is used as an opponent "
                "fallback where possible. Missing opponent inputs stay blank.")
        eligible_in_scope = eligible_players
    else:
        eligible_codes = {TEAM_CODES[t] for t in opponents if t in TEAM_CODES}
        eligible_in_scope = [p for p in eligible_players if p["team"] in eligible_codes]
        st.caption(f"{len(pairs)} games · {len(eligible_in_scope)} eligible players in scope")
    display_six_build_rankings(eligible_in_scope, team_values,
                               data["Player Stats"]["boxes"], metadata, opponents)

st.divider()
st.caption("NBA statistics remain from the uploaded Excel workbook; daily opponent pairings "
           "are from the schedule service. Missing statistics are not guessed or changed to zero. "
           "All displayed player/team base statistics show one decimal place.")