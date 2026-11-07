#!/usr/bin/env python3
"""Orange NBA Viewer — upload an NBA statistics .xlsx when the Streamlit app opens.

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


# Four NEW score tables. Each category uses exactly two 0/3/6 statistics
# and four 0/1/2 statistics. Keys refer to the existing Team Stats headers.
# Display label, Excel field name, point scale maximum.
MATCHUP_SCORE_GROUPS = {
    "Paint Score": [
        ("Paint PTS", "OPP PTS PAINT", 6),
        ("2PTM", "OPP 2PM", 6),
        ("RA FGM", "RA FGM", 2),
        ("RA FGA", "RA FGA", 2),
        ("NRA FGM", "NRA FGM", 2),
        ("NRA FGA", "NRA FGA", 2),
    ],
    "Shooting Score": [
        ("ATB 3PM", "ATB 3PM", 6),
        ("ATB 3PA", "ATB 3PA", 6),
        ("C3PM", "C3PM", 2),
        ("C3PA", "C3PA", 2),
        ("MR FGM", "MR FGM", 2),
        ("MR FGA", "MR FGA", 2),
    ],
    "Scoring Score": [
        ("OPP PTS", "OPP PTS", 6),
        ("OPP FGM", "OPP FGM", 6),
        ("2nd Chance", "OPP PTS 2ND CHANCE", 2),
        ("OREB", "OPP OREB", 2),
        ("FTA", "OPP FTA", 2),
        ("FGA", "OPP FGA", 2),
    ],
    "Fantasy Score": [
        ("Steals", "OPP STL", 6),
        ("Blocks", "OPP BLK", 6),
        ("Off Turnovers", "OPP PTS OFF TOV", 2),
        ("Fastbreak", "OPP PTS FB", 2),
        ("DREB", "OPP DREB", 2),
        ("AST", "OPP AST", 2),
    ],
}


def team_values_for_scoring(sheet_data: dict) -> dict[str, dict]:
    """Flatten the four source boxes, preserving raw values for every team."""
    all_fields = [field for box in sheet_data["boxes"] for field in box["fields"]]
    required_fields = [field for specs in MATCHUP_SCORE_GROUPS.values()
                       for _label, field, _max_points in specs]
    missing = sorted(set(required_fields) - set(all_fields))
    if missing:
        raise ValueError("Team Stats is missing score fields: " + ", ".join(missing))
    if len(all_fields) != len(set(all_fields)):
        raise ValueError("Team Stats includes duplicate statistic headers.")
    return {
        entry["name"]: dict(zip(all_fields,
                                [value for group in entry["groups"] for value in group]))
        for entry in sheet_data["entries"]
    }


def matchup_points(field: str, value: object, maximum: int) -> int | None:
    """Convert an existing 3-tier classification into a weighted score."""
    tier = matchup_tier(field, value)
    if tier is None:
        return None  # Missing is NOT 0 points.
    if maximum == 6:
        return {"Unfavorable": 0, "Neutral": 3, "Favorable": 6}[tier]
    if maximum == 2:
        return {"Unfavorable": 0, "Neutral": 1, "Favorable": 2}[tier]
    raise ValueError(f"Unsupported score maximum: {maximum}")


def build_matchup_score_table(all_teams: dict[str, dict], group_name: str) -> pd.DataFrame:
    """Score all six source stats; preserve partial category totals and coverage."""
    specs = MATCHUP_SCORE_GROUPS[group_name]
    if len(specs) != 6 or len({field for _, field, _ in specs}) != 6:
        raise ValueError(f"{group_name} needs six unique team stats.")
    rows = []
    for team, values in all_teams.items():
        row = {"Team": team}
        for label, source_field, maximum in specs:
            row[label] = matchup_points(source_field, values.get(source_field), maximum)
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
    """Each of the 24 matchup sources counts once; missing values never become 0."""
    specs = [field for group in MATCHUP_SCORE_GROUPS.values() for _, field, _ in group]
    if len(specs) != 24 or len(set(specs)) != 24:
        raise ValueError("Team Matchup Score must use 24 unique source statistics.")
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
        row["Status"] = "Complete" if present == 24 else ("Partial" if present else "Unavailable")
        row["Missing"] = "; ".join(missing)
        rows.append(row)
    return pd.DataFrame(rows)


def point_text_style(value: object, maximum: int) -> str:
    """Color only the score text, never the table cell background."""
    if pd.isna(value):
        return ""
    points = int(value)
    if points == 0:
        return TIER_STYLES["Unfavorable"]
    if points == maximum:
        return TIER_STYLES["Favorable"]
    return TIER_STYLES["Neutral"]


def score_threshold_reference(group_name: str) -> pd.DataFrame:
    rows = []
    for label, source_field, maximum in MATCHUP_SCORE_GROUPS[group_name]:
        unfavorable, neutral, favorable = tier_ranges(source_field)
        medium_points = 3 if maximum == 6 else 1
        rows.append({
            "Score column": label,
            "Excel source": source_field,
            "Unfavorable": f"{unfavorable} → 0",
            "Neutral": f"{neutral} → {medium_points}",
            "Favorable": f"{favorable} → {maximum}",
        })
    return pd.DataFrame(rows)


def show_score_table(frame: pd.DataFrame, group_name: str) -> None:
    """All six score columns, plus partial coverage, with number-only coloring."""
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
    st.caption(f"{len(result)} teams · 6 scoring stats · maximum 20. 'Partial' means some "
               "stats were missing; they are not counted as zero.")
    st.dataframe(styled, hide_index=True, use_container_width=True,
                 height=min(640, 80 + 35 * len(result)))
    with st.expander("See point thresholds and source columns"):
        st.dataframe(score_threshold_reference(group_name),
                     hide_index=True, use_container_width=True)
    st.download_button("Download score table CSV", result.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"orange_nba_{group_name.lower().replace(' ', '_')}.csv",
                       mime="text/csv", key=f"score_{group_name}_csv")


# PLAYER SCORING — ORIGINAL STATISTICS, NO OPPONENT-BASED MULTIPLIERS
# The daily slate still identifies the opponent, but player statistics are never
# multiplied by the opponent's favorable/unfavorable team thresholds.
# The full list also preserves some additional, non-scoring stats in Player Detail.
PLAYER_DETAIL_STATS = (
    "PPG", "2PM", "2PA", "RA FGM", "RA FGA", "NRA FGM", "NRA FGA",
    "3PM", "3PA", "C3PM", "C3PA", "MR FGM", "MR FGA",
    "FTA", "OREB", "DREB", "AST", "STL", "BLK", "FGM", "FGA",
    "%OREB", "%DREB", "%AST", "USG%", "FGM %UAST", "DD2", "TD3",
)
PLAYER_SCORE_CATEGORIES = {
    "Paint": ("2PM", "2PA", "RA FGM", "RA FGA", "NRA FGM", "NRA FGA"),
    "Shooting": ("3PM", "3PA", "C3PM", "C3PA", "MR FGM", "MR FGA"),
    "Attacking": ("USG%", "PPG", "FGM", "FGA", "%OREB", "FGM %UAST"),
    "Fantasy": ("STL", "BLK", "%DREB", "%AST", "DD2", "TD3"),
}
# One source stat, player C3PA, is absent from the strict 56-column NBA workbook.
# It appears as blank in detail and is excluded EQUALLY for every player's core total.
OVERALL_UNAVAILABLE_STATS = {"C3PA"}
OVERALL_SCORE_STATS = tuple(dict.fromkeys(
    field for group in PLAYER_SCORE_CATEGORIES.values()
    for field in group if field not in OVERALL_UNAVAILABLE_STATS
))

# All 225+ cases receive 15; lineup count zero only when a lineup CSV was uploaded.
ROTO_LINEUP_TIERS = [
    (1, 9, 1), (10, 24, 2), (25, 44, 3), (45, 64, 4),
    (65, 84, 5), (85, 99, 6), (100, 115, 7), (116, 130, 8),
    (131, 145, 9), (146, 160, 10), (161, 175, 11),
    (176, 194, 12), (195, 210, 13), (211, 224, 14),
    (225, float("inf"), 15),
]


def roto_points(appearances: object) -> float:
    count = valid_number_or_nan(appearances)
    if pd.isna(count) or count < 0 or int(count) != count:
        return float("nan")
    if count == 0:
        return 0.0
    return float(next(points for low, high, points in ROTO_LINEUP_TIERS if low <= count <= high))


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
    # The fixed 56-column workbook contains 2PM and 3PM, but not an FGM column.
    # Recover the per-game FGM approximation from those displayed values rather
    # than silently dropping the newly requested Attacking FGM statistic.
    if "FGM" not in result or pd.isna(result["FGM"]):
        two = result.get("2PM")
        three = result.get("3PM")
        if pd.notna(two) and pd.notna(three):
            result["FGM"] = round(float(two) + float(three), 1)
    return result, conflicts


def build_player_stat_rows(
    player_data: dict, team_data: dict, team_to_opponent: dict[str, str],
) -> tuple[pd.DataFrame, list[str]]:
    """One original per-game value per player/stat; opponents are informational.

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


def date_in_filename(filename: str) -> date | None:
    """Detect a slate date in filenames from any day; never guess one for CSVs.

    Accepts examples such as NBA_Stats_Nov_07_2026.xlsx, nba_2026-11-07.xlsx,
    20261107-lineups.csv and 11-07-2026-projections.csv. Files without dates
    return None, so a date mismatch is never invented.
    """
    iso = re.search(r"(?<!\d)(20\d{2})[-_]?([01]\d)[-_]?([0-3]\d)(?!\d)", filename)
    if iso:
        try:
            return date(*(int(part) for part in iso.groups()))
        except ValueError:
            pass

    named = re.search(
        r"(?i)(?<![a-z])([a-z]{3,9})[-_ ]+(\d{1,2})[-_ ]+(20\d{2})(?!\d)", filename
    )
    if named:
        value = " ".join(named.groups())
        for fmt in ("%b %d %Y", "%B %d %Y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                pass

    us = re.search(r"(?<!\d)([01]?\d)[-_]([0-3]?\d)[-_](20\d{2})(?!\d)", filename)
    if us:
        try:
            return datetime.strptime(" ".join(us.groups()), "%m %d %Y").date()
        except ValueError:
            pass
    return None


def inferred_game_date(filename: str) -> date:
    """Use the uploaded workbook's slate date, or today's date in U.S. Eastern time."""
    return date_in_filename(filename) or datetime.now(ZoneInfo("America/New_York")).date()


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


def matchup_selection_ui(teams: list[str], upload_name: str) -> list[tuple[str, str]]:
    """Pick the date once for every tab, then auto-pair the complete slate."""
    st.subheader("Game slate")
    st.caption("Game pairings are fetched from ESPN's NBA daily scoreboard feed. All statistics "
               "and player calculations still come exclusively from your uploaded Excel file.")
    col_date, col_mode, col_refresh = st.columns([1.1, 1.8, 0.65])
    with col_date:
        slate_date = st.date_input("Slate date (U.S. Eastern)",
                                   value=inferred_game_date(upload_name),
                                   key=f"selected_slate_date_{upload_name}")
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

        boxes = [
            {"name": (f"Team Box #{i+1}" if sheet == "Team Stats" else f"Player Stats #{i+1}"),
             "fields": headers[stat_start+6*i:stat_start+6*(i+1)]}
            for i in range(expected_boxes)
        ]
        entries = []
        for source in rows[2:]:
            row = (source + [None] * width)[:width]
            if row[0] is None or str(row[0]).strip() == "":
                continue
            entry = {"name": str(row[0]).strip(),
                     "groups": [row[stat_start+6*i:stat_start+6*(i+1)] for i in range(expected_boxes)]}
            if sheet == "Player Stats":
                entry["team"] = str(row[1]).strip() if row[1] is not None else ""
            entries.append(entry)
        if not entries:
            raise ValueError(f'"{sheet}" has no data rows.')
        data[sheet] = {"boxes": boxes, "entries": entries}

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

PROJECTION_LABELS = ["Player Price", "Projected Floor", "Projected Ceiling", "Price per Dollar", "Smash"]
PROJECTION_SYNONYMS = {
    "Player Price": ("player price", "salary", "price", "cost"),
    "Projected Floor": ("projected floor", "proj floor", "floor", "floor points"),
    "Projected Ceiling": ("projected ceiling", "proj ceiling", "ceiling", "ceil"),
    "Price per Dollar": ("price per dollar", "price/dollar", "price per $", "ppd", "value", "p/$", "pts per dollar", "fpts/$", "fpts per dollar"),
    "Smash": ("smash", "smash %", "smash rate", "smash probability"),
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
}


def csv_col_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).casefold())


def find_column(cols: list[str], choices: tuple[str, ...]) -> str | None:
    key_map = {csv_col_key(x): x for x in cols}
    return next((key_map[csv_col_key(word)] for word in choices
                 if csv_col_key(word) in key_map), None)


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


def optional_choice(label: str, columns: list[str], suggested: str | None, key: str) -> str | None:
    options = ["(not provided)"] + columns
    selected = st.selectbox(label, options,
                            index=options.index(suggested) if suggested in options else 0,
                            key=key)
    return None if selected == "(not provided)" else selected


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
            entry = player_index[key]
        else:
            display_name = re.sub(r"\s*\(\d{2,}\)\s*$", "", str(raw)).strip()
            entry = {"name": display_name, "team": code,
                     "groups": [[None] * 6 for _ in range(9)],
                     "_projection_only": True}
        identity = (clean_player_name(entry["name"]), code)
        if identity not in result:
            result[identity] = entry
    return list(result.values()), warnings


def source_csv_inputs(players: list[dict], workbook_filename: str = "",
                      team_names: list[str] | None = None,
                      lineup_file=None, projection_file=None) -> tuple[dict | None, dict, list[str], list[dict], bool]:
    """Return lineup counts, projections, warnings, eligible player rows, CSV loaded.

    A player MUST appear in the selected projections CSV to appear in either
    player-facing section. Projection-only athletes remain visible with missing
    NBA stats; lineup appearances are counted against that eligible roster.
    """
    warnings = []
    eligible_players = []
    projections = {}
    counts = None
    projection_loaded = False
    codes = {TEAM_CODES[t] for t in (team_names or TEAM_CODES) if t in TEAM_CODES}
    with st.expander("CSV data — Optimized Lineups + Projections", expanded=True):
        st.caption("The projections CSV determines which players are shown. Upload it to "
                   "populate Player Stats and Player Matchup Rankings. Players absent "
                   "from the selected projection pool are excluded, including from "
                   "lineup counts. Players in projections but absent from the workbook "
                   "are retained with blank NBA statistics. The lineups CSV is optional "
                   "and supplies Roto points.")
        st.caption("Uploaded optimized lineups: " +
                   (lineup_file.name if lineup_file is not None else "Not uploaded"))
        st.caption("Uploaded Roto projections: " +
                   (projection_file.name if projection_file is not None else "Not uploaded"))
        workbook_date = date_in_filename(workbook_filename)
        for source_name, source_file in (("Optimized Lineups", lineup_file),
                                         ("Projections", projection_file)):
            source_date = date_in_filename(source_file.name) if source_file else None
            if workbook_date is not None and source_date is not None and source_date != workbook_date:
                st.warning(f"Slate date mismatch: {source_name} file is dated "
                           f"{source_date.isoformat()} but NBA workbook is dated "
                           f"{workbook_date.isoformat()}.")
        # Parse the eligibility source FIRST. Do not count optimizer players who
        # aren't in the selected projections pool.
        if projection_file is not None:
            try:
                projection_csv = read_optional_csv(projection_file.getvalue())
                cols = projection_csv.columns.tolist()
                slate_col = find_column(cols, ("SLATE",))
                if slate_col:
                    slate_values = projection_csv[slate_col].fillna("").astype(str).str.strip().str.upper()
                    slate_sizes = slate_values.value_counts()
                    available_slates = [v for v in ("MAIN", "LATE") if v in slate_sizes]
                    if len(available_slates) > 1:
                        st.caption("Available projection pools: " +
                                   ", ".join(f"{g} ({int(slate_sizes[g])})" for g in available_slates))
                        # A Night/late optimizer export goes with the LATE projection pool,
                        # not with the combined MAIN + LATE projection list.
                        lineup_name = lineup_file.name if lineup_file is not None else ""
                        auto_late = bool(re.search(r"(?i)(?:^|[^a-z])(night|late)(?:$|[^a-z])", lineup_name))
                        default_pool = "LATE" if auto_late and "LATE" in available_slates else "All pools"
                        choices = ["All pools"] + available_slates
                        # Each new pair of CSVs gets a fresh default; a prior day's
                        # selected pool must not silently carry over.
                        source_key = re.sub(r"[^a-zA-Z0-9]+", "_",
                                            f"{projection_file.name}_{lineup_name}")[:180]
                        chosen_slate = st.selectbox("Projection slate", choices,
                                                    index=choices.index(default_pool),
                                                    key=f"csv_projection_slate_{source_key}")
                        if auto_late and chosen_slate == "LATE":
                            st.caption("Night lineup export detected: LATE projection pool selected. "
                                       "You may change this if necessary.")
                        if chosen_slate != "All pools":
                            projection_csv = projection_csv[slate_values == chosen_slate]
                st.caption(f"Projection file: {len(projection_csv)} CSV rows selected")
                name_col = optional_choice("Projection player column", cols,
                                           find_column(cols, NAME_SYNONYMS), "csv_projection_name")
                team_col = optional_choice("Projection team column (required for safe matching)",
                                           cols, find_column(cols, TEAM_SYNONYMS), "csv_projection_team")
                field_cols = {
                    label: optional_choice(label, cols, find_column(cols, aliases),
                                           "csv_proj_" + csv_col_key(label))
                    for label, aliases in PROJECTION_SYNONYMS.items()
                }
                eligible_players, roster_warnings = projection_eligible_players(
                    projection_csv, players, name_col, team_col, codes)
                warnings.extend(f"Projection eligibility: {w}" for w in roster_warnings)
                field_mappings = [f"{label} ← {column}" for label, column in field_cols.items() if column]
                if field_mappings:
                    st.caption("Projection columns: " + " · ".join(field_mappings))
                missing_fields = [label for label, column in field_cols.items() if not column]
                if missing_fields:
                    st.warning("Projection columns not mapped: " + ", ".join(missing_fields))
                projections, not_matched = projections_from_csv(
                    projection_csv, eligible_players, name_col, team_col, field_cols)
                warnings.extend(f"Projections: {name}" for name in not_matched)
                projection_loaded = bool(eligible_players and name_col and team_col)
                if projection_loaded:
                    projection_only = sum(p.get("_projection_only", False) for p in eligible_players)
                    st.success(f"Projection-eligible: {len(eligible_players)} players "
                               f"({projection_only} without workbook NBA statistics). "
                               f"Excluded from the {len(players)}-player workbook: "
                               f"{len(players) - (len(eligible_players) - projection_only)} rows.")
                else:
                    st.warning("No eligible players found in the selected projection pool. "
                               "Check the name/team column selections and slate pool.")
            except (ValueError, UnicodeError, OSError, pd.errors.EmptyDataError) as exc:
                st.error(f"Projection CSV error: {exc}")
        if lineup_file is not None:
            if not projection_loaded:
                st.warning("Upload a valid projections CSV first. Players absent from "
                           "projections cannot receive lineup counts or appear in rankings.")
            else:
                try:
                    lineup_csv = read_optional_csv(lineup_file.getvalue())
                    st.caption(f"Optimized Lineups: {len(lineup_csv)} CSV rows")
                    slots = roster_position_columns(lineup_csv.columns.tolist())
                    possible_player = find_column(lineup_csv.columns.tolist(), NAME_SYNONYMS)
                    possible_count = find_column(lineup_csv.columns.tolist(), COUNT_SYNONYMS)
                    default_mode = ("Wide lineup rows" if len(slots) >= 2 else
                                    "Precomputed count" if possible_count and possible_player else
                                    "Long player rows")
                    modes = ["Wide lineup rows", "Long player rows", "Precomputed count"]
                    mode = st.selectbox("Lineup CSV format", modes, index=modes.index(default_mode),
                                        key="csv_lineup_mode")
                    name_col = team_col = id_col = count_col = None
                    used_roster = []
                    cols = lineup_csv.columns.tolist()
                    if mode == "Wide lineup rows":
                        used_roster = st.multiselect("Roster-slot columns (one CSV row = one lineup)",
                                                     options=cols, default=slots, key="csv_roster_cols")
                    else:
                        name_col = optional_choice("Lineup player column", cols,
                                                   possible_player, "csv_lineup_name")
                        team_col = optional_choice("Lineup team column (optional)", cols,
                                                   find_column(cols, TEAM_SYNONYMS), "csv_lineup_team")
                        if mode == "Precomputed count":
                            count_col = optional_choice("Appearance count column", cols,
                                                        possible_count, "csv_lineup_count")
                        else:
                            id_col = optional_choice("Lineup ID column (optional)", cols,
                                                     find_column(cols, LINEUP_ID_SYNONYMS), "csv_lineup_id")
                    counts_observed, not_matched, appearances = count_lineup_appearances(
                        lineup_csv, eligible_players, mode, name_col, team_col, id_col,
                        count_col, used_roster)
                    # An optimizer file from a DIFFERENT slate must not assign a
                    # fabricated zero Roto score to every projected player.
                    if appearances == 0 and len(lineup_csv) > 0:
                        counts = None
                        st.warning("None of the optimized lineup selections match the "
                                   "selected projection pool. Roto points are blank, "
                                   "not zero. Check MAIN vs LATE or the CSV slate date.")
                    else:
                        counts = counts_observed
                        st.success(f"Lineups: {len(counts)} eligible players selected "
                                   f"{appearances} times across the uploaded lineups.")
                    if not_matched:
                        warnings.extend(f"Lineups (not on selected projections or unmatched): {n}"
                                        for n in not_matched)
                except (ValueError, UnicodeError, OSError, pd.errors.EmptyDataError) as exc:
                    st.error(f"Lineup CSV error: {exc}")
        if warnings:
            with st.expander(f"Review {len(warnings)} CSV matching/data warnings"):
                st.write(warnings[:150])
                if len(warnings) > 150:
                    st.caption("Showing the first 150 warnings.")
    return counts, projections, warnings, eligible_players, projection_loaded


def player_metadata(players: list[dict], counts: dict | None, projections: dict) -> pd.DataFrame:
    rows = []
    for entry in players:
        key = entry["name"], entry["team"]
        count = (counts.get(key, 0) if counts is not None else float("nan"))
        row = {"Player": key[0], "Team": key[1],
               "NBA Stats Source": ("Not in workbook" if entry.get("_projection_only")
                                    else "Workbook"),
               "Lineup Appearances": count, "Roto": roto_points(count)}
        row.update(projections.get(key, {label: float("nan") for label in PROJECTION_LABELS}))
        rows.append(row)
    columns = ["Player", "Team", "NBA Stats Source", "Lineup Appearances", "Roto", *PROJECTION_LABELS]
    return pd.DataFrame(rows, columns=columns)


def build_player_overall_leaderboard(long_data: pd.DataFrame,
                                     metadata: pd.DataFrame) -> pd.DataFrame:
    ids = ["Player", "Team", "Opponent"]
    cols = ["Rank", *ids, *PLAYER_SCORE_CATEGORIES, "Roto", "Overall Score",
            "Stats Present", "Score Status", "Missing Stats", "NBA Stats Source",
            "Lineup Appearances", *PROJECTION_LABELS]
    if long_data.empty:
        return pd.DataFrame(columns=cols)
    selected = long_data[long_data["Stat"].isin(OVERALL_SCORE_STATS)].copy()
    if selected.duplicated(ids + ["Stat"]).any():
        raise ValueError("Duplicate player/opponent/stat rows in ranking input.")
    selected["Value"] = selected["Value"].map(valid_number_or_nan)
    wide = selected.pivot(index=ids, columns="Stat", values="Value")
    wide = wide.reindex(columns=list(OVERALL_SCORE_STATS))
    scores = pd.DataFrame(index=wide.index)
    for category, stat_columns in PLAYER_SCORE_CATEGORIES.items():
        available_columns = [col for col in stat_columns if col not in OVERALL_UNAVAILABLE_STATS]
        scores[category] = wide[available_columns].sum(axis=1, min_count=1)
    scores["Stats Present"] = wide.notna().sum(axis=1)
    scores["Missing Stats"] = wide.apply(lambda row: ", ".join(row.index[row.isna()]), axis=1)
    result = scores.reset_index().merge(metadata, on=["Player", "Team"], how="left",
                                        validate="many_to_one")
    # Roto is the supplied lineup-count points (0-15), or missing if no valid CSV.
    columns_to_add = [*PLAYER_SCORE_CATEGORIES, "Roto"]
    result["Overall Score"] = result[columns_to_add].sum(axis=1, min_count=1)
    total_stats = len(OVERALL_SCORE_STATS)
    # No lineup CSV does NOT mean NBA stats are incomplete. A player can have
    # a valid NBA-only score; projections-only players can have a Roto-only score.
    result["Score Status"] = result.apply(
        lambda row: ("Unavailable" if pd.isna(row["Overall Score"]) else
                     "Roto only" if row["Stats Present"] == 0 else
                     "Complete" if row["Stats Present"] == total_stats else "Partial"), axis=1)
    result["Rank"] = result["Overall Score"].rank(ascending=False, method="min").astype("Int64")
    return result[cols].sort_values(["Overall Score", "Player", "Team"],
                                    ascending=[False, True, True], kind="stable",
                                    na_position="last").reset_index(drop=True)


def team_flag_summary(teams: dict[str, dict], selected: list[str]) -> pd.DataFrame:
    rows = []
    for team in selected:
        if team not in teams:
            continue
        values = teams[team]
        positive = []
        negative = []
        neutral = 0
        for stat in TEAM_MATCHUP_THRESHOLDS:
            raw = valid_number_or_nan(values.get(stat))
            tier = matchup_tier(stat, raw)
            if tier == "Favorable":
                positive.append(f"{stat} {raw:.1f}")
            elif tier == "Unfavorable":
                negative.append(f"{stat} {raw:.1f}")
            elif tier == "Neutral":
                neutral += 1
        rows.append({"Team": team, "Positive count": len(positive),
                     "Negative count": len(negative), "Neutral count": neutral,
                     "Positive / Favorable for opponent": ", ".join(positive) or "—",
                     "Negative / Unfavorable for opponent": ", ".join(negative) or "—"})
    return pd.DataFrame(rows).sort_values(["Positive count", "Team"],
                                           ascending=[False, True]).reset_index(drop=True)


def display_player_rankings(long_data: pd.DataFrame, metadata: pd.DataFrame) -> None:
    if long_data.empty:
        st.info("No player/opponent matches available. Check the selected daily slate.")
        return
    leaderboard = build_player_overall_leaderboard(long_data, metadata)
    st.caption("Overall Score adds the original, unadjusted Paint, Shooting, Attacking "
               "and Fantasy statistics plus Roto lineup points when available. No "
               "opponent-based multipliers are applied. Blank or invalid statistics "
               "stay blank, and other valid numbers still contribute to partial scores. "
               "This custom index is not a projected NBA point total or standard "
               "fantasy score; some percentages and shot statistics overlap.")
    st.caption("The opponent is displayed for slate context only and does not affect "
               "player scores. FGM is derived from workbook 2PM + 3PM when needed "
               "and may reflect rounding. Corner 3PA is not in the Excel template "
               "and is excluded equally from overall totals.")
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
            partial = st.checkbox("Show partial / Roto-only scores", value=True, key="leaderboard_partial")
        visible = leaderboard.copy()
        if team_filter != "All teams":
            visible = visible[visible["Team"] == team_filter]
        if search.strip():
            visible = visible[visible["Player"].str.contains(search.strip(), case=False,
                                                            regex=False, na=False)]
        if not partial:
            visible = visible[visible["Score Status"] == "Complete"]
        main_cols = ["Rank", "Player", "Team", "Opponent", *PLAYER_SCORE_CATEGORIES,
                     "Roto", "Overall Score", "Lineup Appearances", "Stats Present",
                     "Score Status", "NBA Stats Source", "Missing Stats", *PROJECTION_LABELS]
        st.dataframe(visible[main_cols], hide_index=True, use_container_width=True,
                     height=min(740, 90 + 33 * max(len(visible), 1)),
                     column_config={**one_decimal_config(visible[main_cols]),
                                    "Rank": st.column_config.NumberColumn("Rank", format="%d"),
                                    "Lineup Appearances": st.column_config.NumberColumn("Lineups", format="%d"),
                                    "Stats Present": st.column_config.NumberColumn("Core Stats", format="%d")})
        complete = (leaderboard["Score Status"] == "Complete").sum()
        partial_count = (leaderboard["Score Status"] == "Partial").sum()
        roto_only = (leaderboard["Score Status"] == "Roto only").sum()
        st.caption(f"{len(leaderboard)} projection-eligible players: {complete} complete / "
                   f"{partial_count} partial / {roto_only} Roto only. "
                   "Missing NBA stats remain blank; available categories still sum. "
                   "Ranks are based on the entire selected slate, not the on-screen filter.")
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
            wide.insert(3, f"{category} Score", wide[list(stats)].sum(axis=1, min_count=1))
            wide = wide.sort_values(f"{category} Score", ascending=False, na_position="last")
            st.caption("Each column shows the original player statistic; the category score sums all available "
                       "statistics in this category. Blank source cells remain blank.")
            st.dataframe(wide, hide_index=True, use_container_width=True,
                         column_config=one_decimal_config(wide))
    with tabs[5]:
        st.caption("Roto = points based on how many optimized lineups contain each player "
                   "(0 if not used in a valid uploaded lineup file; blank if no valid file).")
        roto_table = leaderboard[["Player", "Team", "Opponent", "Lineup Appearances",
                                  "Roto", *PROJECTION_LABELS]].sort_values("Roto", ascending=False,
                                                                          na_position="last")
        st.dataframe(roto_table, hide_index=True, use_container_width=True,
                     column_config=one_decimal_config(roto_table))
        with st.expander("Lineup appearance scoring bands"):
            st.dataframe(pd.DataFrame([
                {"Appearances": "225+" if math.isinf(high) else f"{low}–{high}",
                 "Roto Points": points}
                for low, high, points in ROTO_LINEUP_TIERS
            ]), hide_index=True, use_container_width=True)
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
            st.dataframe(detail[["Stat", "Value"]],
                         hide_index=True, use_container_width=True,
                         column_config={"Value": st.column_config.NumberColumn(format="%.1f")})
    st.download_button("Download original player scoring stats CSV",
                       long_data.to_csv(index=False).encode("utf-8-sig"),
                       file_name="orange_nba_player_scoring_stats.csv", mime="text/csv", key="new_long_download")

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
st.caption("Daily game-slate lookup · Team matchup flags · Four score groups · Original player stats and Roto.")

st.subheader("Upload today's slate files")
st.caption("Upload all three files here, in any order. Excel uses .xlsx; "
           "both DraftKings optimizer and Roto projections use .csv. "
           "These uploaders stay visible even before you upload the workbook.")
workbook_col, optimizer_col, projections_col = st.columns(3)
with workbook_col:
    uploaded = st.file_uploader(
        "1. NBA master workbook (.xlsx)", type=["xlsx"], key="nba_workbook",
        help="The two-sheet Player Stats / Team Stats workbook for this slate date.")
with optimizer_col:
    lineup_upload = st.file_uploader(
        "2. Optimized Lineups (.csv)", type=["csv"], key="optimized_csv",
        help="Eight DraftKings lineup positions per row (PG, SG, SF, PF, C, G, F, UTIL).")
with projections_col:
    projection_upload = st.file_uploader(
        "3. Roto Projections (.csv)", type=["csv"], key="projection_csv",
        help="Player projection rows with PLAYER, TEAM, SALARY, FLOOR, CEIL, FPTS/$, SMASH.")
for source_label, selected_file in (("Workbook", uploaded),
                                    ("Optimized Lineups", lineup_upload),
                                    ("Roto Projections", projection_upload)):
    if selected_file is not None:
        st.caption(f"{source_label} ready: {selected_file.name} "
                   f"({len(selected_file.getvalue()):,} bytes)")
if uploaded is None:
    st.info("Your two CSVs can be selected now. To calculate scores, also upload a "
            "daily NBA master workbook (.xlsx). The selected CSV files will remain "
            "available when the workbook is added.")
    st.stop()
try:
    data = cached_workbook(uploaded.getvalue())
    team_values = team_values_for_scoring(data["Team Stats"])
except (OSError, ValueError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
    st.error(f"Unable to read this NBA workbook: {exc}")
    st.caption("Expected Team Stats: 25 columns; Player Stats: 56 columns; "
               "grouped headers in row 1, column headers in row 2.")
    st.stop()
team_names = data["teamNames"]
players = data["Player Stats"]["entries"]
st.success(f"Workbook loaded: {len(team_names)} teams and {len(players)} players.")

counts, projections, csv_warnings, eligible_players, has_projection_pool = source_csv_inputs(
    players, uploaded.name, team_names,
    lineup_file=lineup_upload, projection_file=projection_upload)
metadata = player_metadata(eligible_players, counts, projections)
# The selected projection file is authoritative for *both* player sections.
# Keep the team sections and original uploaded workbook unchanged.
visible_player_data = {**data["Player Stats"], "entries": eligible_players}
selected_slate_pairs = matchup_selection_ui(team_names, uploaded.name)
slate_pair_labels = {game_label(a, b): (a, b) for a, b in selected_slate_pairs}
slate_team_names = [team for pair in selected_slate_pairs for team in pair]

team_main, score_main, player_main, player_adjust_main = st.tabs(
    ["Team Stats", "Matchup Score", "Player Stats", "Player Rankings"]
)

with team_main:
    st.subheader("Team Stats — all five boxes")
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

    st.markdown("### Team highlights — positive and negative matchup flags")
    highlights = team_flag_summary(team_values, selected_teams)
    if highlights.empty:
        st.info("No teams selected.")
    else:
        st.dataframe(highlights, hide_index=True, use_container_width=True,
                     height=min(620, 85 + 36 * len(highlights)),
                     column_config={"Positive / Favorable for opponent": st.column_config.TextColumn(width="large"),
                                    "Negative / Unfavorable for opponent": st.column_config.TextColumn(width="large")})
        st.caption("All favorable and unfavorable source stats appear in the text columns. "
                   "Neutral source stats are counted but not called out. Missing source "
                   "values are not automatically treated as negative flags.")
    for index, box in enumerate(data["Team Stats"]["boxes"]):
        st.divider()
        st.markdown(f"### Team Box {index + 1} — six stats")
        frame = six_stat_table(data["Team Stats"], "Team Stats", index)
        frame = frame[frame["Team"].isin(selected_teams)]
        show_six_stat_tab(frame, section="teams", box_name=box["name"])
    st.divider()
    st.markdown("### Team Box 5 — reserved for additional outlier statistics")
    st.info("Box 5 is ready as a new section, but its six source statistics and outlier "
            "definitions have not been provided. No invented numbers or thresholds are shown.")
    blank_box = pd.DataFrame([{"Team": team, **{f"Future Stat {i}": None for i in range(1, 7)}}
                              for team in selected_teams])
    if not blank_box.empty:
        st.dataframe(blank_box, hide_index=True, use_container_width=True)

with score_main:
    st.subheader("Matchup Score")
    st.caption("Four scoring categories (Paint, Shooting, Scoring, Fantasy), each out of 20. "
               "All 24 source statistics are included exactly once. Missing inputs generate "
               "a Partial score rather than making the entire team's score disappear.")
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
        all_scores = build_overall_score_table(team_values)
        visible_scores = all_scores[all_scores["Team"].isin(scoring_teams)]
        st.markdown("**Overall score / 80**")
        visible_scores = visible_scores.sort_values("Overall / 80", ascending=False,
                                                    na_position="last", kind="stable")
        st.dataframe(visible_scores, hide_index=True, use_container_width=True,
                     column_config=one_decimal_config(visible_scores))
        st.caption("Team totals are based on their own defense allowed — the higher the score, "
                   "the more favorable that matchup is for the opposing offense. "
                   "Coverage = Stats Present out of 24. Partial totals can be understated.")
        score_tabs = st.tabs(list(MATCHUP_SCORE_GROUPS))
        for category, tab in zip(MATCHUP_SCORE_GROUPS, score_tabs):
            with tab:
                st.markdown(f"**{category} — six source stats, 20 points maximum**")
                category_df = build_matchup_score_table(team_values, category)
                category_df = category_df[category_df["Team"].isin(scoring_teams)]
                show_score_table(category_df, category)
                with st.expander("Actual defense-allowed values behind these points"):
                    source_df = pd.DataFrame([
                        {"Team": t, **{label: valid_number_or_nan(team_values[t].get(field))
                                       for label, field, _ in MATCHUP_SCORE_GROUPS[category]}}
                        for t in scoring_teams if t in team_values
                    ])
                    if not source_df.empty:
                        styled = source_df.style
                        for label, source, _ in MATCHUP_SCORE_GROUPS[category]:
                            styled = styled.map(lambda value, field=source: style_team_cell(value, field),
                                                subset=[label])
                        st.dataframe(styled.format({c: "{:.1f}" for c in source_df.columns[1:]},
                                                  na_rep="—"), hide_index=True, use_container_width=True)
    except ValueError as exc:
        st.error(f"Matchup Score unavailable: {exc}")

with player_main:
    st.subheader("Player Stats")
    st.caption("Your nine original groups of six remain unchanged. Only players listed in "
               "the selected Roto projections CSV are shown here. Blank NBA source "
               "statistics stay blank; other statistics are unaffected.")
    if not has_projection_pool:
        st.info("Upload the Roto projections CSV above to show eligible players. "
                "Workbook-only and lineup-only players are not displayed.")
    else:
        st.caption(f"Showing {len(eligible_players)} projection-eligible players; "
                   f"{sum(p.get('_projection_only', False) for p in eligible_players)} "
                   "have no NBA workbook stat row.")
    st.markdown("### Lineup counts and projection fields")
    a, b = st.columns([1, 2])
    with a:
        chosen_team = st.selectbox("Team filter", ["All teams"] + sorted(set(p["team"] for p in eligible_players)),
                                   key="player_filter_team")
    with b:
        search = st.text_input("Find player", placeholder="Search name", key="player_search")
    csv_view = metadata.copy()
    if chosen_team != "All teams":
        csv_view = csv_view[csv_view["Team"] == chosen_team]
    if search.strip():
        csv_view = csv_view[csv_view["Player"].str.contains(search.strip(), case=False,
                                                             regex=False, na=False)]
    st.dataframe(csv_view, hide_index=True, use_container_width=True,
                 column_config=one_decimal_config(csv_view))
    if counts is None:
        st.caption("No valid Optimized Lineups CSV: lineup appearances and Roto points are blank.")
    if not projections:
        st.caption("No matched Projections CSV: price, floor, ceiling, price per dollar and Smash stay blank.")
    st.divider()
    player_tabs = st.tabs([box["name"] for box in data["Player Stats"]["boxes"]])
    for index, (tab, box) in enumerate(zip(player_tabs, data["Player Stats"]["boxes"])):
        with tab:
            frame = six_stat_table(visible_player_data, "Player Stats", index)
            if chosen_team != "All teams":
                frame = frame[frame["Team"] == chosen_team]
            if search.strip():
                frame = frame[frame["Player"].str.contains(search.strip(), case=False,
                                                           regex=False, na=False)]
            show_six_stat_tab(frame, section="players", box_name=box["name"])

with player_adjust_main:
    st.subheader("Player Rankings — Original Stats + Roto")
    st.caption("Only players in the selected projections pool can appear here. "
               "Scores use original NBA statistics with no matchup multipliers. "
               "The daily slate supplies opponent labels and optional game filtering. "
               "Blank NBA values are skipped, never substituted with zero; valid "
               "statistics and uploaded Roto lineup points still add to the score.")
    if not has_projection_pool:
        st.info("Upload your Roto projections CSV above to enable player rankings.")
    game_options = ["Entire selected slate", *slate_pair_labels]
    selected_game = st.selectbox("Game scope", game_options, key="player_adjustment_selection_mode")
    pairs = (selected_slate_pairs if selected_game == "Entire selected slate"
             else [slate_pair_labels[selected_game]])
    opponents = {}
    for away, home in pairs:
        opponents[away] = home
        opponents[home] = away
    # Pairings control game-scope filtering when available, but they are NOT a
    # prerequisite for scoring: unadjusted NBA stats can be ranked without ESPN.
    if not opponents:
        st.info("No matchups selected. Player rankings still work with original NBA "
                "statistics; opponent labels stay blank until a slate is selected.")
        eligible_in_scope = eligible_players
    else:
        eligible_codes = {TEAM_CODES[t] for t in opponents if t in TEAM_CODES}
        eligible_in_scope = [p for p in eligible_players if p["team"] in eligible_codes]
        st.caption(f"{len(pairs)} games / {len(opponents)} teams / "
                   f"{len(eligible_in_scope)} of {len(eligible_players)} "
                   "projection-eligible players in the selected games.")
    scoped_player_data = {**visible_player_data, "entries": eligible_in_scope}
    try:
        long_data, conflicts = build_player_stat_rows(scoped_player_data, team_values, opponents)
        if conflicts:
            st.warning("Conflicting duplicate player statistics were excluded from "
                       "the affected calculation. Other valid stats still receive scores.")
            with st.expander("See conflicting fields"):
                st.write(conflicts)
        display_player_rankings(long_data, metadata)
    except (ValueError, KeyError) as exc:
        st.error(f"Player scoring could not be calculated: {exc}")

st.divider()
st.caption("NBA statistics remain from the uploaded Excel workbook; daily opponent pairings "
           "are from the schedule service. Missing statistics are not guessed or changed to zero. "
           "All displayed player/team base statistics show one decimal place.")
