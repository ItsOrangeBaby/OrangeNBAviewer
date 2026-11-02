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
    "OPP 2PM": (25, 31, None),
    "OPP PTS PAINT": (43, 54, None),
    "OPP PTS 2ND CHANCE": (14, 19, None),
    "OPP OREB": (13, 18, None),
    "RA FGM": (14, 20, None),
    "OPP FTA": (22, 29, None),

    "OPP FGM": (39, 44, 40.94),
    "OPP AST": (24, 29, 26.09),
    "NRA FGM": (6, 10, 7.78),
    "MR FGM": (3, 5, 3.56),
    "C3PM": (3, 5, 3.46),
    "ATB 3PM": (8, 11, 9.47),

    "OPP FGA": (86, 93, 89.13),
    "OPP PTS OFF TOV": (17, 22, 19.09),
    "OPP PTS FB": (13, 19, 15.18),
    "OPP DREB": (29, 36, 32.40),
    "OPP STL": (7, 11, 8.78),
    "OPP BLK": (5, 7, 5.17),

    "OPP PTS": (106, 124, 114.68),
    "RA FGA": (21, 30, 25.15),
    "NRA FGA": (14, 20, 17.46),
    "MR FGA": (6, 11, 8.89),
    "C3PA": (9, 12, 10.08),
    "ATB 3PA": (24, 31, 27.54),
}

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
    neutral_start, favorable_start, _average = thresholds
    if number >= favorable_start:
        return "Favorable"
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
        low_end, high_start, average = TEAM_MATCHUP_THRESHOLDS[field]
        row = {
            "Statistic": field,
            "Unfavorable": f"< {low_end}",
            "Neutral": f"{low_end} to < {high_start}",
            "Favorable": f"≥ {high_start}",
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


def build_matchup_score_table(
    all_teams: dict[str, dict], group_name: str
) -> pd.DataFrame:
    """Team + six numeric point columns + a 0–20 category total."""
    specs = MATCHUP_SCORE_GROUPS[group_name]
    rows = []
    for team, values in all_teams.items():
        row = {"Team": team}
        for label, source_field, maximum in specs:
            row[label] = matchup_points(source_field, values.get(source_field), maximum)
        scores = [row[label] for label, _field, _maximum in specs]
        row["Total / 20"] = sum(scores) if all(v is not None for v in scores) else None
        rows.append(row)
    frame = pd.DataFrame(rows)
    for col in frame.columns[1:]:
        frame[col] = pd.to_numeric(frame[col], errors="coerce")
    return frame


def build_overall_score_table(all_teams: dict[str, dict]) -> pd.DataFrame:
    """The four independent category totals and the 0–80 combined total."""
    scores_by_group = {group: build_matchup_score_table(all_teams, group).set_index("Team")
                       for group in MATCHUP_SCORE_GROUPS}
    rows = []
    for team in all_teams:
        row = {"Team": team}
        for group, frame in scores_by_group.items():
            row[group] = frame.at[team, "Total / 20"]
        values = [row[group] for group in MATCHUP_SCORE_GROUPS]
        row["Overall / 80"] = sum(values) if all(pd.notna(x) for x in values) else None
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
        neutral_start, favorable_start, _average = TEAM_MATCHUP_THRESHOLDS[source_field]
        medium_points = 3 if maximum == 6 else 1
        rows.append({
            "Score column": label,
            "Excel source": source_field,
            "Unfavorable": f"< {neutral_start} → 0",
            "Neutral": f"{neutral_start} to < {favorable_start} → {medium_points}",
            "Favorable": f"≥ {favorable_start} → {maximum}",
        })
    return pd.DataFrame(rows)


def show_score_table(frame: pd.DataFrame, group_name: str) -> None:
    """Sortable category table, with six individually colored numeric scores."""
    if frame.empty:
        st.info("No matching teams. Select another matchup or team filter.")
        return
    columns = list(frame.columns[1:])
    left, right = st.columns([2, 1])
    with left:
        sort_column = st.selectbox("Sort by", ["Total / 20"] + columns[:-1] + ["Team"],
                                   key=f"score_{group_name}_sort")
    with right:
        sort_direction = st.selectbox("Order", ["Highest first", "Lowest first"],
                                      key=f"score_{group_name}_direction")
    result = frame.sort_values(sort_column,
                               ascending=(sort_direction == "Lowest first"),
                               kind="stable", na_position="last").reset_index(drop=True)
    styled = result.style
    specs = MATCHUP_SCORE_GROUPS[group_name]
    for label, _source_field, maximum in specs:
        styled = styled.map(lambda value, m=maximum: point_text_style(value, m), subset=[label])
    styled = styled.map(lambda value: "font-weight: 800;" if pd.notna(value) else "",
                        subset=["Total / 20"])
    styled = styled.format({col: "{:.0f}" for col in columns}, na_rep="—")
    st.caption(f"{len(result)} teams · 6 scored statistics · maximum 20 points. "
               "Click a column heading to sort.")
    st.dataframe(styled, hide_index=True, use_container_width=True,
                 height=min(640, 80 + 35 * len(result)))
    with st.expander("See point thresholds and source columns"):
        st.dataframe(score_threshold_reference(group_name),
                     hide_index=True, use_container_width=True)
    st.download_button("Download score table CSV", result.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"orange_nba_{group_name.lower().replace(' ', '_')}.csv",
                       mime="text/csv", key=f"score_{group_name}_csv")


# PLAYER MATCHUP ADJUSTMENTS
# The team's *opponent-allowed* statistic determines the multiplier applied
# to each offensive player's own per-game statistic. The list of defensive
# sources is ordered, and each distinct source is applied ONCE.
# If multiple sources affect one player statistic, MULTIPLY their multipliers.
# Unfavorable x0.5, Neutral x1.0, Favorable x1.5.
PLAYER_MULTIPLIERS = {"Unfavorable": 0.5, "Neutral": 1.0, "Favorable": 1.5}
PLAYER_MATCHUP_LINKS = {
    "PPG": ("OPP PTS",),
    "2PM": ("OPP 2PM", "OPP AST"),
    "2PA": ("OPP 2PM",),
    "RA FGM": ("OPP PTS PAINT", "RA FGM", "RA FGA"),
    "RA FGA": ("OPP PTS PAINT", "RA FGM", "RA FGA"),
    "NRA FGM": ("OPP PTS PAINT", "NRA FGA"),
    "NRA FGA": ("OPP PTS PAINT", "NRA FGA"),
    "3PM": ("OPP AST", "ATB 3PA"),
    "3PA": ("ATB 3PA",),
    "C3PM": ("C3PA",),
    # C3PA isn't included in the 56-column original Player Stats sheet.
    # We still show this requested adjustment, with its baseline left blank.
    "C3PA": ("C3PA",),
    "MR FGM": ("MR FGA",),
    "MR FGA": ("MR FGA",),
    "FTA": ("OPP FTA",),
    "OREB": ("OPP OREB", "OPP PTS 2ND CHANCE"),
    "DREB": ("OPP DREB", "OPP PTS FB"),
    "AST": ("OPP AST",),
    "STL": ("OPP STL", "OPP PTS OFF TOV"),
    "BLK": ("OPP BLK",),
}
PLAYER_MATCHUP_CATEGORIES = {
    "Paint & 2PT": ("2PM", "2PA", "RA FGM", "RA FGA", "NRA FGM", "NRA FGA"),
    "Shooting": ("3PM", "3PA", "C3PM", "C3PA", "MR FGM", "MR FGA"),
    "Other": ("PPG", "FTA", "OREB", "DREB", "AST", "STL", "BLK"),
}


def player_raw_stat_values(entry: dict, boxes: list[dict]) -> tuple[dict, list[str]]:
    """Flatten the nine player groups, checking duplicate column names."""
    result = {}
    conflicts = []
    for box, group in zip(boxes, entry["groups"]):
        for label, raw in zip(box["fields"], group):
            number = pd.to_numeric(raw, errors="coerce")
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
    return result, conflicts


def compute_player_adjustment(stat: str, baseline: object, opponent_values: dict) -> dict:
    """One independent factor per referenced defensive metric, multiplied together."""
    factors = []
    for defense_field in dict.fromkeys(PLAYER_MATCHUP_LINKS[stat]):
        allowed = pd.to_numeric(opponent_values.get(defense_field), errors="coerce")
        tier = matchup_tier(defense_field, allowed)
        factor = PLAYER_MULTIPLIERS[tier] if tier is not None else None
        factors.append((defense_field, allowed, tier, factor))
    available = all(factor is not None for _field, _val, _tier, factor in factors)
    multiplier = math.prod(x[3] for x in factors) if available else float("nan")
    base = pd.to_numeric(baseline, errors="coerce")
    adjusted = base * multiplier if pd.notna(base) and available else float("nan")
    return {"Baseline": base, "Multiplier": multiplier, "Adjusted": adjusted,
            "Factors": factors}


def build_player_matchup_rows(
    player_data: dict, team_data: dict, team_to_opponent: dict[str, str],
) -> tuple[pd.DataFrame, list[str]]:
    """Build long-format comparisons, never matching a player to their own defense."""
    team_by_code = {TEAM_CODES[name]: name for name in team_data if name in TEAM_CODES}
    rows = []
    warnings = set()
    for entry in player_data["entries"]:
        team_name = team_by_code.get(entry["team"])
        opponent = team_to_opponent.get(team_name)
        if opponent is None or opponent == team_name or opponent not in team_data:
            continue
        player_values, conflicts = player_raw_stat_values(entry, player_data["boxes"])
        if conflicts:
            warnings.add(f'{entry["name"]} ({entry["team"]}): conflicting values for '
                         + ", ".join(conflicts))
        for stat in PLAYER_MATCHUP_LINKS:
            result = compute_player_adjustment(stat, player_values.get(stat), team_data[opponent])
            factor_parts = []
            for defense_field, defensive_value, tier, factor in result["Factors"]:
                if tier is None:
                    factor_parts.append(f"{defense_field}: missing")
                else:
                    factor_parts.append(f"{defense_field} {float(defensive_value):.1f} "
                                        f"({tier}, x{factor:g})")
            rows.append({
                "Player": entry["name"], "Team": entry["team"],
                "Opponent": opponent, "Stat": stat,
                "Baseline": result["Baseline"], "Multiplier": result["Multiplier"],
                "Adjusted": result["Adjusted"],
                "Defense factors": " × ".join(factor_parts),
            })
    columns = ["Player", "Team", "Opponent", "Stat", "Baseline", "Multiplier",
               "Adjusted", "Defense factors"]
    return pd.DataFrame(rows, columns=columns), sorted(warnings)


# Player C3PA is not included in the original 56-column Excel player schema.
# Exclude it explicitly from the overall score for ALL players so no one is
# rewarded for having less missing data; retain its blank individual adjustment.
OVERALL_UNAVAILABLE_STATS = {"C3PA"}
OVERALL_SCORE_STATS = tuple(
    stat for stat in PLAYER_MATCHUP_LINKS if stat not in OVERALL_UNAVAILABLE_STATS
)


def build_player_overall_leaderboard(long_data: pd.DataFrame) -> pd.DataFrame:
    """Rank complete players by the unweighted sum of their 18 adjusted stats.

    Each calculated stat counts once, including overlapping volume categories,
    exactly as requested. Any missing stat other than permanently unavailable
    C3PA makes an individual player's combined score *unavailable*, not zero.
    """
    columns = ["Rank", "Player", "Team", "Opponent", *PLAYER_MATCHUP_CATEGORIES,
               "Overall Score", "Stats Present", "Missing Stats"]
    if long_data.empty:
        return pd.DataFrame(columns=columns)

    ids = ["Player", "Team", "Opponent"]
    selected = long_data[long_data["Stat"].isin(OVERALL_SCORE_STATS)].copy()
    if selected.duplicated(ids + ["Stat"]).any():
        raise ValueError("Duplicate player / opponent / adjusted-stat rows in ranking input.")

    wide = selected.pivot(index=ids, columns="Stat", values="Adjusted")
    # Reindex to ensure missing columns do not accidentally disappear.
    wide = wide.reindex(columns=list(OVERALL_SCORE_STATS))
    scores = pd.DataFrame(index=wide.index)
    for category, stats in PLAYER_MATCHUP_CATEGORIES.items():
        fields = [stat for stat in stats if stat in OVERALL_SCORE_STATS]
        # Require all expected stats: no zero-filling of missing source values.
        scores[category] = wide[fields].sum(axis=1, min_count=len(fields))

    scores["Overall Score"] = scores[list(PLAYER_MATCHUP_CATEGORIES)].sum(
        axis=1, min_count=len(PLAYER_MATCHUP_CATEGORIES)
    )
    scores["Stats Present"] = wide.notna().sum(axis=1)
    scores["Missing Stats"] = wide.apply(
        lambda row: ", ".join(row.index[row.isna()]), axis=1
    )
    # Equal scores share the same rank; ranking is across all selected matchups,
    # before any optional player/team filters in the UI.
    scores["Rank"] = scores["Overall Score"].rank(
        ascending=False, method="min"
    ).astype("Int64")
    result = scores.reset_index()
    return result[columns].sort_values(
        ["Overall Score", "Player", "Team"],
        ascending=[False, True, True], kind="stable", na_position="last"
    ).reset_index(drop=True)


def display_player_adjustments(long_data: pd.DataFrame) -> None:
    """Overall player leaderboard, stat-specific rankings and factor breakdown."""
    if long_data.empty:
        st.info("No matching players/opponent defenses found. Change the matchup selection.")
        return

    leaderboard = build_player_overall_leaderboard(long_data)
    overall_tab, ranking, breakdown = st.tabs([
        "Overall player leaderboard", "Individual stat rankings", "Player breakdown"
    ])

    with overall_tab:
        required = len(OVERALL_SCORE_STATS)
        ranked_count = int(leaderboard["Overall Score"].notna().sum())
        st.markdown("**Overall Player Matchup Score — ranked highest to lowest**")
        st.caption(
            "Overall Score = Paint & 2PT + Shooting + Other. Each category is the SUM "
            "of its matchup-adjusted player statistics, using every opponent multiplier "
            "you specified. This is a custom raw-stat index, not projected points or "
            "a standard fantasy score; points, makes, and attempts overlap."
        )
        st.caption(
            f"{ranked_count} of {len(leaderboard)} players have all {required} required "
            "adjusted stats. Player C3PA is excluded for everyone because the current "
            "Excel template does not provide it."
        )
        col_team, col_find, col_incomplete = st.columns([1, 2, 1])
        with col_team:
            ranking_team = st.selectbox(
                "Team", ["All teams"] + sorted(leaderboard["Team"].unique().tolist()),
                key="overall_player_team_filter"
            )
        with col_find:
            player_search = st.text_input(
                "Find player", placeholder="Search name", key="overall_player_search"
            )
        with col_incomplete:
            show_incomplete = st.checkbox(
                "Show incomplete", value=True, key="overall_player_show_incomplete",
                help="Incomplete player scores are blank and excluded from ranks, not zero."
            )

        visible = leaderboard.copy()
        if ranking_team != "All teams":
            visible = visible[visible["Team"] == ranking_team]
        if player_search.strip():
            visible = visible[
                visible["Player"].str.contains(player_search.strip(), case=False,
                                               regex=False, na=False)
            ]
        if not show_incomplete:
            visible = visible[visible["Overall Score"].notna()]

        st.dataframe(
            visible, hide_index=True, use_container_width=True,
            height=min(680, 85 + 35 * max(len(visible), 1)),
            column_config={
                "Rank": st.column_config.NumberColumn("Rank", format="%d"),
                **{
                    category: st.column_config.NumberColumn(category, format="%.1f")
                    for category in PLAYER_MATCHUP_CATEGORIES
                },
                "Overall Score": st.column_config.NumberColumn(
                    "Overall Score", format="%.1f"
                ),
                "Stats Present": st.column_config.NumberColumn(
                    f"Stats / {required}", format="%d"
                ),
                "Missing Stats": st.column_config.TextColumn("Missing Stats"),
            },
        )
        st.caption("Ranks are calculated across all players in the selected matchups; "
                   "filtering by team or name does not change the original rank. "
                   "Click a heading to sort the table.")
        st.download_button(
            "Download overall player leaderboard CSV",
            leaderboard.to_csv(index=False).encode("utf-8-sig"),
            file_name="orange_nba_overall_player_matchup_ranking.csv",
            mime="text/csv", key="overall_player_matchup_csv"
        )
        incomplete_count = len(leaderboard) - ranked_count
        if incomplete_count:
            st.warning(
                f"{incomplete_count} players have additional missing source statistics. "
                "Their overall scores and ranks stay blank until those values are available."
            )

    with ranking:
        a, b = st.columns([2, 1])
        with a:
            category = st.selectbox("Statistic category", list(PLAYER_MATCHUP_CATEGORIES),
                                    key="player_adjustment_category")
        with b:
            ranking_stat = st.selectbox("Player stat", list(PLAYER_MATCHUP_CATEGORIES[category]),
                                        key=f"player_adjust_rank_{category}")
        sub = long_data[long_data["Stat"] == ranking_stat].copy()
        sub = sub.sort_values("Adjusted", ascending=False, na_position="last", kind="stable")
        st.caption(f"{len(sub)} players · Adjusted = original stat × product of all defensive "
                   "tier multipliers. Click headings to sort.")
        st.dataframe(sub[["Player", "Team", "Opponent", "Baseline", "Multiplier", "Adjusted"]],
                     hide_index=True, use_container_width=True,
                     column_config={
                         "Baseline": st.column_config.NumberColumn("Original", format="%.1f"),
                         "Multiplier": st.column_config.NumberColumn("Total x", format="%.3f"),
                         "Adjusted": st.column_config.NumberColumn("Adjusted", format="%.1f"),
                     })
        if ranking_stat == "C3PA":
            st.warning("The original Player Stats workbook does not include player corner-three "
                       "attempts (C3PA). Adjusted C3PA stays blank; this is not a zero.")

    with breakdown:
        player_options = (long_data[["Player", "Team"]]
                          .drop_duplicates().sort_values(["Team", "Player"]))
        labels = [f"{r.Player} ({r.Team})" for r in player_options.itertuples(index=False)]
        player_choice = st.selectbox("Player", labels, key="player_adjustment_detail_player")
        selected_name, selected_team = player_options.iloc[labels.index(player_choice)].tolist()
        rows = long_data[(long_data["Player"] == selected_name) &
                         (long_data["Team"] == selected_team)].copy()
        total_row = leaderboard[(leaderboard["Player"] == selected_name) &
                                (leaderboard["Team"] == selected_team)].iloc[0]
        score_value = total_row["Overall Score"]
        rank_value = total_row["Rank"]
        top_a, top_b, top_c = st.columns(3)
        with top_a:
            st.metric("Overall Score", f"{score_value:.1f}" if pd.notna(score_value) else "—")
        with top_b:
            st.metric("Overall Rank", f"#{int(rank_value)}" if pd.notna(rank_value) else "—")
        with top_c:
            st.metric("Stats available", f"{int(total_row['Stats Present'])}/{len(OVERALL_SCORE_STATS)}")
        st.caption(f"{selected_name} · Original per-game values vs. "
                   f"{rows.iloc[0]['Opponent']} defense")
        st.dataframe(rows[["Stat", "Baseline", "Multiplier", "Adjusted", "Defense factors"]],
                     hide_index=True, use_container_width=True,
                     column_config={
                         "Baseline": st.column_config.NumberColumn("Original", format="%.1f"),
                         "Multiplier": st.column_config.NumberColumn("Total x", format="%.3f"),
                         "Adjusted": st.column_config.NumberColumn("Adjusted", format="%.1f"),
                     })
    st.download_button("Download all player adjustments CSV",
                       long_data.to_csv(index=False).encode("utf-8-sig"),
                       file_name="orange_nba_player_matchup_adjustments.csv",
                       mime="text/csv", key="player_adjustments_csv")


PREFERRED_MATCHUPS = [
    ("Brooklyn Nets", "Boston Celtics"),
    ("Detroit Pistons", "New York Knicks"),
    ("Sacramento Kings", "San Antonio Spurs"),
    ("Portland Trail Blazers", "Los Angeles Lakers"),
]

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

    frame = pd.DataFrame(rows)
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
            st.caption("Inclusive cutoff rule: '25–30' means 25.0 up to, "
                       "but not including, 31.0. Decimal values follow the same rule.")
    else:
        st.dataframe(sorted_frame, hide_index=True, use_container_width=True,
                     height=min(640, 80 + 35 * len(sorted_frame)),
                     column_config=one_decimal_config(sorted_frame))
    st.download_button("Download this group's CSV", sorted_frame.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"orange_nba_{section}_{box_name.lower().replace(' ', '_').replace('#','')}.csv",
                       mime="text/csv", key=f"{section}_{box_name}_csv")


st.set_page_config(page_title="Orange NBA Viewer", page_icon="🏀", layout="wide",
                   initial_sidebar_state="collapsed")
st.markdown("""
<style>
.block-container {max-width:1450px; padding-top:1.5rem; padding-bottom:3rem;}
.nba-kicker {color:#f97316;font-size:.78rem;font-weight:800;letter-spacing:.12em;text-transform:uppercase;}
/* Keep separate box tabs readable, including on a phone. */
.stTabs [data-baseweb="tab-list"] {gap:.25rem;}
.stTabs [data-baseweb="tab"] {font-weight:650;white-space:nowrap;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="nba-kicker">Orange NBA Viewer</div>', unsafe_allow_html=True)
st.title("NBA Stats Explorer")
st.caption("Upload your NBA Excel workbook. Upload an NBA workbook to explore Team Stats, Matchup Score, Player Stats, and Player Matchup Adjustments.")

uploaded = st.file_uploader("Upload your NBA Excel workbook (.xlsx)", type=["xlsx"],
                            help="Requires the original Player Stats and Team Stats worksheet format.")
if uploaded is None:
    st.info("Upload an .xlsx file to browse four Team Stats tabs, four Matchup Score tabs, nine Player Stats tabs, and Player Matchup Adjustments. "
            "You do not need to put the workbook in GitHub.")
    st.stop()

try:
    data = cached_workbook(uploaded.getvalue())
except (OSError, ValueError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
    st.error(f"Unable to read this workbook: {exc}")
    st.caption("Expected: Team Stats (25 columns) and Player Stats (56 columns), "
               "grouped titles in row 1, column headers in row 2.")
    st.stop()

team_names = data["teamNames"]
team_entries = {entry["name"]: entry for entry in data["Team Stats"]["entries"]}
players = data["Player Stats"]["entries"]
st.success(f"Loaded {uploaded.name} — {len(team_names)} teams and {len(players)} players.")

team_main, score_main, player_main, player_adjust_main = st.tabs(
    ["Team Stats", "Matchup Score", "Player Stats", "Player Matchup Adjustments"]
)

with team_main:
    st.subheader("Team Stats")
    st.caption("Select a team or matchup, then switch among the four six-stat boxes. "
               "Highlights describe how favorable that DEFENSE is for the opposing offense "
               "(more allowed = more favorable).")
    st.markdown("**Matchup highlights:** 🟥 Unfavorable · 🟨 Neutral · 🟩 Favorable "
                "· Missing values are uncolored")

    presets = [(a, b) for a, b in PREFERRED_MATCHUPS if a in team_entries and b in team_entries]
    team_options = ["All teams"] + [f"{a} vs {b}" for a, b in presets] + ["Choose specific teams"]
    team_filter = st.selectbox("Show teams", team_options, key="team_filter")
    if team_filter == "Choose specific teams":
        selected_teams = st.multiselect("Teams to display", team_names, default=team_names,
                                        key="selected_team_names")
    elif team_filter == "All teams":
        selected_teams = team_names
    else:
        selected_teams = list(presets[team_options.index(team_filter) - 1])

    team_boxes = data["Team Stats"]["boxes"]
    team_box_tabs = st.tabs([box["name"] for box in team_boxes])
    for box_index, (tab, box) in enumerate(zip(team_box_tabs, team_boxes)):
        with tab:
            st.markdown(f"**{box['name']}** — 6 statistics")
            frame = six_stat_table(data["Team Stats"], "Team Stats", box_index)
            frame = frame[frame["Team"].isin(selected_teams)]
            show_six_stat_tab(frame, section="teams", box_name=box["name"])

with score_main:
    st.subheader("Matchup Score")
    st.caption("Scores measure how favorable a team's DEFENSE is for the opposing offense. "
               "More opponent production allowed = more points for the opponent. "
               "Every category has two 0/3/6 stats and four 0/1/2 stats.")
    st.markdown("**Score colors:** 🔴 0 = Unfavorable · 🟡 1 or 3 = Neutral "
                "· 🟢 2 or 6 = Favorable · Missing = —")

    # Independent selection, so Matchup Score can be used without touching Team Stats.
    score_presets = [(a, b) for a, b in PREFERRED_MATCHUPS
                     if a in team_entries and b in team_entries]
    score_options = (["All teams"] + [f"{a} vs {b}" for a, b in score_presets]
                     + ["Choose specific teams"])
    score_filter = st.selectbox("Show teams / matchup", score_options, key="score_team_filter")
    if score_filter == "Choose specific teams":
        scoring_teams = st.multiselect("Teams to score", team_names,
                                       default=team_names, key="scoring_team_names")
    elif score_filter == "All teams":
        scoring_teams = team_names
    else:
        scoring_teams = list(score_presets[score_options.index(score_filter) - 1])
        offense_a, offense_b = scoring_teams
        st.caption(f"Matchup interpretation: {offense_a}'s row measures opportunities "
                   f"for {offense_b}; {offense_b}'s row measures opportunities "
                   f"for {offense_a}.")

    try:
        scoring_values = team_values_for_scoring(data["Team Stats"])
    except ValueError as exc:
        st.error(str(exc))
    else:
        all_scores = build_overall_score_table(scoring_values)
        all_scores = all_scores[all_scores["Team"].isin(scoring_teams)]
        with st.expander("Overall Matchup Scores (out of 80)", expanded=True):
            if all_scores.empty:
                st.info("No teams selected.")
            else:
                overall_sorted = all_scores.sort_values("Overall / 80", ascending=False,
                                                       kind="stable", na_position="last")
                overall_styled = overall_sorted.style.format(
                    {col: "{:.0f}" for col in overall_sorted.columns[1:]}, na_rep="—")
                st.dataframe(overall_styled, hide_index=True, use_container_width=True)
                st.caption("A missing statistic makes its category and overall score blank "
                           "rather than counting the missing value as zero.")

        score_tabs = st.tabs(list(MATCHUP_SCORE_GROUPS))
        for group_name, tab in zip(MATCHUP_SCORE_GROUPS, score_tabs):
            with tab:
                st.markdown(f"**{group_name}** — maximum 20 points")
                score_frame = build_matchup_score_table(scoring_values, group_name)
                score_frame = score_frame[score_frame["Team"].isin(scoring_teams)]
                show_score_table(score_frame, group_name)

                with st.expander("See actual defensive stats behind these scores"):
                    source_frame = pd.DataFrame([
                        {"Team": name, **{
                            label: pd.to_numeric(scoring_values[name].get(field), errors="coerce")
                            for label, field, _maximum in MATCHUP_SCORE_GROUPS[group_name]
                        }}
                        for name in scoring_teams if name in scoring_values
                    ])
                    if source_frame.empty:
                        st.info("No teams selected.")
                    else:
                        styled_source = source_frame.style
                        for label, field, _maximum in MATCHUP_SCORE_GROUPS[group_name]:
                            styled_source = styled_source.map(
                                lambda value, f=field: style_team_cell(value, f), subset=[label])
                        styled_source = styled_source.format(
                            {label: "{:.1f}" for label in source_frame.columns[1:]}, na_rep="—")
                        st.dataframe(styled_source, hide_index=True, use_container_width=True)


with player_main:
    st.subheader("Player Stats")
    st.caption("Each player box shows Player, Team, and exactly six statistics. "
               "Sort each box independently from highest to lowest or lowest to highest.")
    codes = sorted({p["team"] for p in players if p["team"]})
    filter_col, search_col = st.columns(2)
    with filter_col:
        chosen_team = st.selectbox("Filter by team", ["All teams"] + codes,
                                   key="player_filter_team")
    with search_col:
        search = st.text_input("Search player", placeholder="e.g. Brunson", key="player_search")

    player_boxes = data["Player Stats"]["boxes"]
    player_box_tabs = st.tabs([box["name"] for box in player_boxes])
    for box_index, (tab, box) in enumerate(zip(player_box_tabs, player_boxes)):
        with tab:
            st.markdown(f"**{box['name']}** — 6 statistics")
            frame = six_stat_table(data["Player Stats"], "Player Stats", box_index)
            if chosen_team != "All teams":
                frame = frame[frame["Team"] == chosen_team]
            if search.strip():
                frame = frame[frame["Player"].str.contains(search.strip(), case=False,
                                                          regex=False, na=False)]
            show_six_stat_tab(frame, section="players", box_name=box["name"])

with player_adjust_main:
    st.subheader("Player Matchup Adjustments")
    st.caption("Match players with the defense they are facing, then multiply each applicable "
               "defensive multiplier. Unfavorable = 0.5×, Neutral = 1.0×, Favorable = 1.5×. "
               "These are heuristic adjustments, not predictions or guaranteed box scores.")

    available_pairs = [(a, b) for a, b in PREFERRED_MATCHUPS
                       if a in team_entries and b in team_entries]
    matchup_options = (["All four preset matchups"] if available_pairs else []) + [
        "Choose an offensive team"]
    selection_mode = st.selectbox("Matchup selection", matchup_options,
                                  key="player_adjustment_selection_mode")
    if selection_mode == "All four preset matchups":
        team_to_opponent = {}
        for first, second in available_pairs:
            team_to_opponent[first] = second
            team_to_opponent[second] = first
        st.caption("Opponents assigned from the four previously specified games; "
                   "only pairs present in the uploaded workbook are included.")
    else:
        offensive_team = st.selectbox("Players on team", team_names,
                                      key="player_adjust_offense")
        opponents = [name for name in team_names if name != offensive_team]
        known_opponent = next((second if first == offensive_team else first
                               for first, second in available_pairs
                               if offensive_team in (first, second)), None)
        opponent_options = ["Choose opponent defense"] + opponents
        chosen_opponent = st.selectbox(
            "Opponent defense", opponent_options,
            index=opponent_options.index(known_opponent) if known_opponent in opponents else 0,
            key=f"player_adjust_defense_{offensive_team}",
        )
        team_to_opponent = ({offensive_team: chosen_opponent}
                            if chosen_opponent != "Choose opponent defense" else {})
        if not team_to_opponent:
            st.info("Select the opposing defense to calculate player adjustments.")

    try:
        all_defenses = team_values_for_scoring(data["Team Stats"])
        adjusted_rows, import_warnings = build_player_matchup_rows(
            data["Player Stats"], all_defenses, team_to_opponent
        )
    except (ValueError, KeyError) as exc:
        st.error(f"Player matchup calculation unavailable: {exc}")
    else:
        if import_warnings:
            st.warning("One or more duplicated player statistics disagree between boxes. "
                       "Those values are left blank instead of selecting one arbitrarily.")
            with st.expander("See data conflicts"):
                st.write(import_warnings)
        display_player_adjustments(adjusted_rows)
        with st.expander("See the team-to-player multiplier rules"):
            st.dataframe(pd.DataFrame([
                {"Player stat": stat, "Opponent defense factors (multiply)": " × ".join(fields)}
                for stat, fields in PLAYER_MATCHUP_LINKS.items()
            ]), hide_index=True, use_container_width=True)
            st.caption("Any missing player baseline or required defensive input results in a "
                       "blank adjusted value, not 0. Neutral factors still multiply as 1.0×. "
                       "The original Player Stats worksheet lacks player C3PA, so that "
                       "adjustment remains unavailable without additional source data.")


st.divider()
st.caption("Statistics are taken only from the workbook you upload. Missing source values "
           "stay blank, not zero. The six statistics in every box retain their exact original order.")
