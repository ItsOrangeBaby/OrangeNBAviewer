#!/usr/bin/env python3
"""Orange NBA Viewer — upload an NBA statistics .xlsx when the Streamlit app opens.

Deploy this file as app.py on Streamlit Community Cloud.
Run locally: streamlit run app.py

Does not require any workbook to be stored in GitHub. The file you upload stays
in the current Streamlit session; it is parsed in memory and is not saved.
"""

from __future__ import annotations

import html
from collections import Counter

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


def esc(value: object) -> str:
    return html.escape(str(value), quote=True)


def displayed(value: object) -> str:
    if value is None or value == "" or value == "-":
        return "—"
    return str(value)


def stat_card(box: dict, left: list, right: list | None = None,
              left_label: str = "Value", right_label: str = "") -> None:
    """Show exactly the six columns in one original NBA workbook box."""
    if len(box["fields"]) != 6 or len(left) != 6 or (right is not None and len(right) != 6):
        raise ValueError("Each statistical box must have exactly six fields")
    comparison = right is not None
    css_class = "" if comparison else " solo"
    pieces = [f'<section class="nba-box"><div class="nba-box-title">{esc(box["name"])}'
              '<span class="nba-box-meta">6 stats</span></div>']
    if comparison:
        pieces.append('<div class="nba-row heading"><div class="nba-label">Statistic</div>'
                      f'<div class="nba-val">{esc(left_label)}</div>'
                      f'<div class="nba-val">{esc(right_label)}</div></div>')
    for i, label in enumerate(box["fields"]):
        pieces.append(f'<div class="nba-row{css_class}"><div class="nba-label">{esc(label)}</div>'
                      f'<div class="nba-val">{esc(displayed(left[i]))}</div>')
        if comparison:
            pieces.append(f'<div class="nba-val alt">{esc(displayed(right[i]))}</div>')
        pieces.append("</div>")
    pieces.append("</section>")
    st.markdown("".join(pieces), unsafe_allow_html=True)


def box_selection(boxes: list, selected: str):
    return [(i, box) for i, box in enumerate(boxes) if selected == "All boxes" or box["name"] == selected]




def stat_columns(boxes: list[dict]) -> list[str]:
    """Create unique, readable labels while preserving the original Excel order.

    Player Stats contains a 3PA field in both boxes 1 and 9. Dataframe columns
    must be unique, so only repeated headings receive a box suffix.
    """
    names = [field for box in boxes for field in box["fields"]]
    counts = Counter(names)
    labels = []
    for number, box in enumerate(boxes, 1):
        for field in box["fields"]:
            labels.append(f"{field} · Box {number}" if counts[field] > 1 else field)
    return labels


def combined_table(sheet_data: dict, sheet_name: str) -> pd.DataFrame:
    """One sortable table per worksheet; all statistics remain numeric."""
    boxes = sheet_data["boxes"]
    headers = stat_columns(boxes)
    is_player = sheet_name == "Player Stats"
    rows = []
    for entry in sheet_data["entries"]:
        record = {"Player": entry["name"], "Team": entry["team"]} if is_player else {"Team": entry["name"]}
        values = [value for group in entry["groups"] for value in group]
        for label, value in zip(headers, values):
            record[label] = value
        rows.append(record)
    frame = pd.DataFrame(rows)
    for name in headers:
        # Missing NBA values (including '-') become NaN, never an invented 0.
        frame[name] = pd.to_numeric(frame[name], errors="coerce")
    return frame


def sortable_grid(frame: pd.DataFrame, *, section: str) -> None:
    """Display every statistic in one horizontal-scroll table with explicit sorting."""
    if frame.empty:
        st.info("No matching rows.")
        return
    first_col, second_col = st.columns([3, 1])
    with first_col:
        sort_field = st.selectbox("Sort by", list(frame.columns), key=f"{section}_sort")
    with second_col:
        direction = st.selectbox("Order", ["Highest first", "Lowest first"],
                                 index=1 if sort_field in ("Team", "Player") else 0,
                                 key=f"{section}_order")
    ascending = direction == "Lowest first"
    ordered = frame.sort_values(sort_field, ascending=ascending,
                                kind="stable", na_position="last").reset_index(drop=True)
    st.caption(f"{len(ordered)} rows · {len(ordered.columns)} columns · Scroll sideways to see every statistic. "
               "You can also tap a column heading to sort interactively.")
    st.dataframe(ordered, hide_index=True, use_container_width=True, height=min(660, 80 + 35 * len(ordered)))
    st.download_button("Download this sorted table as CSV", ordered.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"orange_nba_{section}_sorted.csv", mime="text/csv", key=f"{section}_csv")


st.set_page_config(page_title="Orange NBA Viewer", page_icon="🏀", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
.block-container {max-width:1450px; padding-top:1.5rem; padding-bottom:3rem;}
.nba-kicker {color:#f97316;font-size:.78rem;font-weight:800;letter-spacing:.12em;text-transform:uppercase;}
.nba-box {border:1px solid #56607455;border-radius:14px;overflow:hidden;margin:.35rem 0 1rem;}
.nba-box-title {padding:.85rem 1rem;background:#f9731615;font-weight:750;font-size:1.02rem;}
.nba-box-meta {opacity:.6;font-size:.73rem;font-weight:500;margin-left:.7rem;}
.nba-row {display:grid;grid-template-columns:minmax(0,1.5fr) minmax(0,.7fr) minmax(0,.7fr);align-items:center;gap:.5rem;padding:.7rem 1rem;border-bottom:1px solid #88888820;}
.nba-row:last-child {border-bottom:0;}
.nba-row.solo {grid-template-columns:minmax(0,1.5fr) minmax(0,.7fr);}
.nba-label {font-size:.88rem;overflow-wrap:anywhere;}
.nba-val {text-align:right;font-size:.98rem;font-weight:700;font-variant-numeric:tabular-nums;}
.nba-val.alt {color:#f97316;}
.nba-row.heading {background:#88888812;font-size:.78rem;opacity:.8;}
@media(max-width:640px){.nba-row{padding:.65rem .6rem;gap:.3rem}.nba-label{font-size:.76rem}.nba-val{font-size:.85rem}.nba-box-title{padding:.75rem .6rem}}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="nba-kicker">Orange NBA Viewer</div>', unsafe_allow_html=True)
st.title("NBA Stats Explorer")
st.caption("Upload an NBA Excel workbook. Team Stats opens first, with all stats in one sortable table.")

uploaded = st.file_uploader("Upload your NBA Excel workbook (.xlsx)", type=["xlsx"],
                            help="Requires the original Player Stats and Team Stats worksheet format.")
if uploaded is None:
    st.info("Upload an .xlsx file to see the complete sortable team and player tables. No data file is needed in GitHub.")
    st.stop()

try:
    data = cached_workbook(uploaded.getvalue())
except (OSError, ValueError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
    st.error(f"Unable to read this workbook: {exc}")
    st.caption("Expected: Team Stats (25 columns) and Player Stats (56 columns), grouped titles in row 1, column headers in row 2.")
    st.stop()

team_names = data["teamNames"]
team_entries = {entry["name"]: entry for entry in data["Team Stats"]["entries"]}
players = data["Player Stats"]["entries"]
team_codes = {TEAM_CODES.get(name, name): name for name in team_names}
team_frame = combined_table(data["Team Stats"], "Team Stats")
player_frame = combined_table(data["Player Stats"], "Player Stats")
st.success(f"Loaded {uploaded.name} — {len(team_names)} teams and {len(players)} players.")

team_tab, player_tab = st.tabs(["Team Stats", "Player Stats"])

with team_tab:
    st.subheader("Team Stats — All 24 statistics")
    st.caption("The figures are defensive stats **allowed to opponents**, not the team's own offensive output. "
               "Higher allowed values indicate more potential opportunity for the opposing offense. "
               "Point thresholds have not been applied.")
    presets = [(a, b) for a, b in PREFERRED_MATCHUPS if a in team_entries and b in team_entries]
    team_options = ["All teams"] + [f"{a} vs {b}" for a, b in presets] + ["Choose specific teams"]
    team_filter = st.selectbox("Show teams", team_options, key="team_filter")
    if team_filter == "Choose specific teams":
        selected_teams = st.multiselect("Teams to display", team_names, default=team_names, key="specific_teams")
    elif team_filter == "All teams":
        selected_teams = team_names
    else:
        selected_teams = list(presets[team_options.index(team_filter) - 1])
    visible_teams = team_frame[team_frame["Team"].isin(selected_teams)]
    sortable_grid(visible_teams, section="teams")

    with st.expander("Optional: show the original six-stat boxes"):
        if not selected_teams:
            st.info("Choose at least one team.")
        else:
            detail_teams = st.multiselect("Team boxes to compare", selected_teams,
                                          default=selected_teams[:2], max_selections=2,
                                          key="detail_teams")
            detail_group = st.selectbox("Six-stat group", ["All boxes"] +
                                        [box["name"] for box in data["Team Stats"]["boxes"]],
                                        key="detail_team_group")
            if len(detail_teams) in (1, 2):
                first = team_entries[detail_teams[0]]
                second = team_entries[detail_teams[1]] if len(detail_teams) == 2 else None
                for index, box in box_selection(data["Team Stats"]["boxes"], detail_group):
                    stat_card(box, first["groups"][index],
                              second["groups"][index] if second is not None else None,
                              TEAM_CODES.get(first["name"], first["name"]),
                              TEAM_CODES.get(second["name"], second["name"]) if second else "")

with player_tab:
    st.subheader("Player Stats — All 54 statistics")
    codes = sorted({p["team"] for p in players if p["team"]})
    left_filter, right_filter = st.columns(2)
    with left_filter:
        chosen_team = st.selectbox("Filter by team", ["All teams"] + codes, key="player_filter_team")
    with right_filter:
        search = st.text_input("Search player", placeholder="e.g. Brunson", key="player_search")

    filtered = player_frame.copy()
    if chosen_team != "All teams":
        filtered = filtered[filtered["Team"] == chosen_team]
    if search.strip():
        filtered = filtered[filtered["Player"].str.contains(search.strip(), case=False, regex=False, na=False)]
    sortable_grid(filtered, section="players")

    with st.expander("Optional: show the original six-stat boxes for one player"):
        matching_players = [p for p in players if
                            (chosen_team == "All teams" or p["team"] == chosen_team) and
                            (search.strip().casefold() in p["name"].casefold())]
        if not matching_players:
            st.info("No players matched your search.")
        else:
            index = st.selectbox("Choose player", range(len(matching_players)),
                                 format_func=lambda i: f'{matching_players[i]["name"]} · {matching_players[i]["team"]}',
                                 key="detail_player")
            detail_group = st.selectbox("Six-stat group", ["All boxes"] +
                                        [box["name"] for box in data["Player Stats"]["boxes"]],
                                        key="detail_player_group")
            entry = matching_players[index]
            for box_index, box in box_selection(data["Player Stats"]["boxes"], detail_group):
                stat_card(box, entry["groups"][box_index])

st.divider()
st.caption("Statistics come only from your uploaded workbook. Missing source values display as blank in tables (not zero). "
           "The columns follow your original box order; repeated names such as 3PA have box suffixes for clarity.")
