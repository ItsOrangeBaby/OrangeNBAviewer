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
    st.dataframe(sorted_frame, hide_index=True, use_container_width=True,
                 height=min(640, 80 + 35 * len(sorted_frame)))
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
st.caption("Upload your NBA Excel workbook. Team Stats opens first, and each tab has exactly six statistics.")

uploaded = st.file_uploader("Upload your NBA Excel workbook (.xlsx)", type=["xlsx"],
                            help="Requires the original Player Stats and Team Stats worksheet format.")
if uploaded is None:
    st.info("Upload an .xlsx file to browse four Team Stats tabs and nine Player Stats tabs. "
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

team_main, player_main = st.tabs(["Team Stats", "Player Stats"])

with team_main:
    st.subheader("Team Stats")
    st.caption("Select a team or matchup, then switch among the four six-stat boxes. "
               "Higher values allowed by a defense indicate a potentially more favorable offensive matchup. "
               "Scoring thresholds have not been added.")

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

st.divider()
st.caption("Statistics are taken only from the workbook you upload. Missing source values "
           "stay blank, not zero. The six statistics in every box retain their exact original order.")
