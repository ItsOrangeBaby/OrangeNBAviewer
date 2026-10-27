#!/usr/bin/env python3
"""Orange NBA Viewer — upload an NBA statistics .xlsx when the Streamlit app opens.

Deploy this file as app.py on Streamlit Community Cloud.
Run locally: streamlit run app.py

Does not require any workbook to be stored in GitHub. The file you upload stays
in the current Streamlit session; it is parsed in memory and is not saved.
"""

from __future__ import annotations

import html
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


st.set_page_config(page_title="Orange NBA Viewer", page_icon="🏀", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
.block-container {max-width:1180px; padding-top:1.5rem; padding-bottom:3rem;}
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
st.caption("Upload a new Excel workbook whenever you want to explore different NBA stats. Team Stats opens first.")

# Crucially, NO hardcoded workbook file is read from GitHub or from the local folder.
uploaded = st.file_uploader("Upload your NBA Excel workbook (.xlsx)", type=["xlsx"],
                            help="Workbook must have Player Stats and Team Stats tabs using the original six-stat group layout.")
if uploaded is None:
    st.info("Upload your .xlsx file above to see Team Stats and Player Stats. No workbook needs to be committed to GitHub.")
    st.stop()

try:
    data = cached_workbook(uploaded.getvalue())
except (OSError, ValueError, KeyError, IndexError, zipfile.BadZipFile, ET.ParseError) as exc:
    st.error(f"Unable to read this workbook: {exc}")
    st.caption("Expected: tabs named 'Team Stats' (25 columns) and 'Player Stats' (56 columns), with group headers on row 1 and column headers on row 2.")
    st.stop()

team_entries = {t["name"]: t for t in data["Team Stats"]["entries"]}
team_names = data["teamNames"]
players = data["Player Stats"]["entries"]
team_codes = {TEAM_CODES.get(name, name): name for name in team_names}

st.success(f"Loaded {uploaded.name} — {len(team_names)} teams and {len(players)} players.")

# Tabs are ordered intentionally: Team Stats before Player Stats.
team_tab, player_tab = st.tabs(["Team Stats", "Player Stats"])
with team_tab:
    st.subheader("Team defensive allowances")
    st.caption("Higher values allowed suggest more opportunity for the opponent. Scoring thresholds have not been added yet.")

    presets = [(a, b) for a, b in PREFERRED_MATCHUPS if a in team_entries and b in team_entries]
    if len(team_names) > 1:
        matchup_choices = [f"{a} vs {b}" for a, b in presets] + ["Custom team comparison"]
        chosen_preset = st.selectbox("Choose a matchup", matchup_choices)
        if chosen_preset == "Custom team comparison":
            first_name = st.selectbox("First team's defense", team_names, index=0)
            second_candidates = [t for t in team_names if t != first_name]
            second_name = st.selectbox("Second team's defense", second_candidates)
        else:
            first_name, second_name = presets[matchup_choices.index(chosen_preset)]
        left_ui, right_ui = st.columns(2)
        with left_ui:
            mode = st.radio("View", ["Compare both defenses", "Single defense"], horizontal=True)
        with right_ui:
            group_choice = st.selectbox("Six-stat group", ["All boxes"] + [b["name"] for b in data["Team Stats"]["boxes"]])
        if mode == "Single defense":
            chosen = st.selectbox("Defense to inspect", [first_name, second_name])
            st.caption(f"Opponent offensive opportunity against {chosen}'s defense")
            for idx, box in box_selection(data["Team Stats"]["boxes"], group_choice):
                stat_card(box, team_entries[chosen]["groups"][idx])
        else:
            st.caption(f"The values below are allowed by each team's defense — not scored by its own offense.")
            for idx, box in box_selection(data["Team Stats"]["boxes"], group_choice):
                stat_card(box, team_entries[first_name]["groups"][idx],
                          team_entries[second_name]["groups"][idx],
                          TEAM_CODES.get(first_name, first_name), TEAM_CODES.get(second_name, second_name))
    else:
        st.caption(f"Defensive stats allowed by {team_names[0]}")
        for idx, box in enumerate(data["Team Stats"]["boxes"]):
            stat_card(box, team_entries[team_names[0]]["groups"][idx])

with player_tab:
    st.subheader("Player statistics")
    codes = sorted({p["team"] for p in players if p["team"]})
    left_ui, right_ui = st.columns(2)
    with left_ui:
        chosen_team = st.selectbox("Filter by team", ["All teams"] + codes)
    with right_ui:
        query = st.text_input("Search player name", placeholder="e.g. Brunson")
    filtered = [p for p in players if (chosen_team == "All teams" or p["team"] == chosen_team)
                and query.strip().casefold() in p["name"].casefold()]
    st.caption(f"{len(filtered)} matching player(s) / {len(players)} total")
    if not filtered:
        st.info("No players matched your search.")
    else:
        chosen_index = st.selectbox("Choose player", list(range(len(filtered))),
                                    format_func=lambda i: f'{filtered[i]["name"]} · {filtered[i]["team"]}')
        group_choice = st.selectbox("Six-stat group", ["All boxes"] +
                                    [b["name"] for b in data["Player Stats"]["boxes"]], key="player_group")
        entry = filtered[chosen_index]
        st.caption(f'{entry["name"]} · {team_codes.get(entry["team"], entry["team"])}')
        for idx, box in box_selection(data["Player Stats"]["boxes"], group_choice):
            stat_card(box, entry["groups"][idx])

st.divider()
st.caption("All statistics come from your uploaded workbook. A dash (—) means missing source data, not zero. Uploading another .xlsx replaces the current view.")
