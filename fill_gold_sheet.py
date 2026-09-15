#!/usr/bin/env python3
"""
Fill Gold Sheet 2.0 (Internal) PDF from:
  - Builder Accounts Salesforce Report (SFXL / .xlsx)
  - Prior Gold Sheet (PGS / .xlsx), e.g. GS - Shea Homes 2026
  - Gold Sheet 2.0 template - Internal Version (fillable PDF)

Does not modify source files. Writes a filled PDF (+ optional JSON preview).

Examples:
  python fill_gold_sheet.py --inspect-pdf "Gold sheet 2.0 template -Internal Version.pdf"
  python fill_gold_sheet.py --inspect-pgs "GS - Shea Homes 2026.xlsx"
  python fill_gold_sheet.py ^
      --sfxl "Builder Accounts Salesforce Report.xlsx" ^
      --pgs "GS - Shea Homes 2026.xlsx" ^
      --pdf "Gold sheet 2.0 template -Internal Version.pdf" ^
      --out "GS - Shea Homes 2026.pdf" ^
      --dry-run
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from typing import Any, Iterable, Optional

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, TextStringObject


# ---------------------------------------------------------------------------
# Label aliases (prior gold sheets vary in wording / punctuation)
# ---------------------------------------------------------------------------

LABEL_ALIASES: dict[str, list[str]] = {
    "Account Name": [
        "Account Name",
        "Builder Name",
        "Company Name",
    ],
    "Builder Rank": ["Builder Rank", "Builder Ranking", "Rank"],
    "Annual Closings": ["Annual Closings", "Closings", "Annual Closing"],
    "National Average Selling Price": [
        "National Average Selling Price",
        "National ASP",
        "National Average Selling Price (ASP)",
        "ASP",
        "National Average Selling Price:",
    ],
    "Spec Home %": [
        "Spec Home %",
        "Spec Home%",
        "Spec %",
        "Spec Percent",
        "Spec Home Percent",
        "Spec Homes %",
    ],
    "Rebate Structure": [
        "Rebate Structure:",
        "Rebate Structure",
    ],
    "Design Center Program": [
        "Design Center Program",
        "Model Home Design Center Program",
        "Model Home / Design Center Program",
        "Model Home Design Center",
    ],
    "2026 Pricing Action Status": [
        "2026 Pricing Action Status",
        "Pricing Action Status",
        "Pricing Actions",
        "2026 Pricing Actions",
    ],
    "National Pricing": [
        "National Pricing:",
        "National Pricing",
    ],
    "Special Terms": ["Special Terms", "Special Term"],
    "Builder Strategy": [
        "Builder's Strategy",
        "Builders Strategy",
        "Builder Strategy",
        "Builder's Strategies",
        "Growth Strategy",
        "Growth Strategies",
        "Builder's",
    ],
    "Operating Divisions": [
        "Operating Divisions",
        "Divisions",
        "Operating Division",
        "Total number of Divisions",
        "Total Number of Divisions",
        "Total Divisions",
    ],
    "National Account Manager": [
        "National Account Manager",
        "NAM",
        "National Account Mgr",
    ],
}

# Some PGS fields should only populate from the cells adjacent to their own label.
# If the label row has no value, do not fall through to a later unrelated label/value pair.
STRICT_ADJACENT_VALUE_LOGICALS = frozenset({"Spec Home %"})

# These PGS fields are commonly stored as vertically merged label blocks with
# one or more value rows to the right. Prefer bounded row extraction over
# loose neighbor/fallback lookup so we do not spill into adjacent sections.
BLOCK_VALUE_LOGICALS = frozenset(
    {
        "Rebate Structure",
        "Design Center Program",
        "2026 Pricing Action Status",
        "National Pricing",
        "Special Terms",
        "Builder Strategy",
    }
)

# Uniform appearance for populated fields (Helvetica — primary form font in the 2.0 template).
UNIFORM_FIELD_FONT = "Helv"
BOLD_FIELD_FONT = "HeBo"  # Helvetica-Bold — available in template /DR
UNIFORM_FIELD_SIZE = 10
UNIFORM_FIELD_SIZE_LARGE = 12
UNIFORM_FIELD_SIZE_STAT = 14  # Builder_Rank, Annual_Closings, National_ASP, Spec_Percent (+2pt)
DA_BLACK_10 = f"/{UNIFORM_FIELD_FONT} {UNIFORM_FIELD_SIZE} Tf 0 g"
DA_BLACK_12 = f"/{UNIFORM_FIELD_FONT} {UNIFORM_FIELD_SIZE_LARGE} Tf 0 g"
DA_BLACK_14 = f"/{UNIFORM_FIELD_FONT} {UNIFORM_FIELD_SIZE_STAT} Tf 0 g"
DA_WHITE_10 = f"/{UNIFORM_FIELD_FONT} {UNIFORM_FIELD_SIZE} Tf 1 1 1 rg"
DA_WHITE_12 = f"/{UNIFORM_FIELD_FONT} {UNIFORM_FIELD_SIZE_LARGE} Tf 1 1 1 rg"

# White text on dark header bar (match template /DA).
WHITE_FIELD_DA: dict[str, str] = {
    "Headquarters": DA_WHITE_10,
    "Website_URL": DA_WHITE_10,
    "Market_Facing_Position": DA_WHITE_10,
    "NAM_Name": DA_WHITE_10,
    "NAM_Email": DA_WHITE_10,
    "Date": DA_WHITE_10,
}

GROWTH_STRAT_LABELS = frozenset(
    {"Growth_Strat_1", "Growth_Strat_2", "Growth_Strat_3", "Growth_Strat_4"}
)

# Stat row fields at 14pt black (+2pt from prior 12pt).
LARGE_TEXT_LOGICALS = frozenset(
    {
        "Builder_Rank",
        "Annual_Closings",
        "National_ASP",
        "Spec_Percent",
    }
)

# PDF AcroForm field name candidates (template naming has drifted slightly)
PDF_FIELD_CANDIDATES: dict[str, list[str]] = {
    "Builder_Name": ["Builder_Name", "Builder Name"],
    "Headquarters": ["Headquarters", "HQ", "Headquaters"],
    "Website_URL": ["Website_URL", "Website URL", "Website"],
    "Market_Facing_Position": [
        "Market_Facing_Position",
        "Market Facing Position",
        "Market_Facing_Positon",
    ],
    "Date": ["Date"],
    "NAM_Name": ["NAM_Name", "NAM Name"],
    "NAM_Email": ["NAM_Email", "NAM Email"],
    "Builder_Rank": ["Builder_Rank", "Builder Rank"],
    "Annual_Closings": ["Annual_Closings", "Annual Closings"],
    "National_ASP": ["National_ASP", "National ASP"],
    "Spec_Percent": ["Spec_Percent", "Spec %", "Spec_Percentt"],
    "Rebate_Structure": ["Rebate_Structure", "Rebate Structure"],
    "Model_Home_Design_Center_Program": [
        "Model_Home_Design_Center_Program",
        "Model Home Design Center Program",
    ],
    "Pricing_Actions": ["Pricing_Actions", "Pricing Actions"],
    "National_Pricing": ["National_Pricing", "National Pricing"],
    "Special_Terms": ["Special_Terms", "Special Terms"],
    "GrowthStrat1_Text": [
        "GrowthStrat1_Text",
        "GrowthStrat1_Text1",
        "Text3",
    ],
    "GrowthStrat2_Text": [
        "GrowthStrat2_Text",
        "GrowthStrat2_Text2",
        "Text444",
    ],
    "GrowthStrat3_Text": [
        "GrowthStrat3_Text",
        "GrowthStrat3_Text3",
        "Text333",
    ],
    "GrowthStrat4_Text": [
        "GrowthStrat4_Text",
        "GrowthStrat4_Text4",
        "Text555",
    ],
    "Growth_Strat_1": ["Growth_Strat_1", "Growth Strat 1", "Growth_Strat 1"],
    "Growth_Strat_2": ["Growth_Strat_2", "Growth Strat 2", "Growth_Strat 2"],
    "Growth_Strat_3": ["Growth_Strat_3", "Growth Strat 3", "Growth_Strat 3"],
    "Growth_Strat_4": ["Growth_Strat_4", "Growth Strat 4", "Growth Strat 4"],
    "Divisions_1": ["Divisions_1", "Divisions"],
    "Divisions_2": ["Divisions_2", "Divisions 2"],
    "Divisions_3": ["Divisions_3", "Divisions 3"],
    "Divisions_4": ["Divisions_4", "Divisions 4"],
    "Total_D": ["Total_D", "Total D", "Total_D "],
}

SFXL_COLUMN_ALIASES: dict[str, list[str]] = {
    # Prefer exact header names only — loose aliases like "Account" wrongly hit "Account Owner".
    "Account Name": ["Account Name", "Builder Name"],
    "Physical State/Province": [
        "Physical State/Province",
        "Physical State / Province",
        "Billing State/Province",
        "Shipping State/Province",
        "State/Province",
        "Physical State",
    ],
    "Website": ["Website", "Website URL", "Web Site"],
    "Type": ["Type", "Account Type"],
    "Last Modified Date": [
        "Last Modified Date",
        "LastModifiedDate",
    ],
}

# Account Name filter: Corporate / Corp as whole words (not "Corpus").
CORPORATE_NAME_RE = re.compile(r"(?i)\bcorp(?:orate)?\b")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def norm(text: Any) -> str:
    if text is None:
        return ""
    s = str(text).replace("\xa0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def norm_key(text: Any) -> str:
    s = norm(text).lower()
    s = s.replace("’", "'").replace("‘", "'")
    s = re.sub(r"[:：]\s*$", "", s)
    s = re.sub(r"[^a-z0-9%+/' ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def cell_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, date):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return norm(value)


def split_multiline(text: str) -> list[str]:
    if not text:
        return []
    parts = re.split(r"[\r\n]+|•|\u2022|\u25cf", text)
    out: list[str] = []
    for p in parts:
        p = norm(p)
        p = re.sub(r"^[\-\*\u2013\u2014]+\s*", "", p)
        p = re.sub(r"^\(?\d{1,2}\)?[.)]\s*", "", p)
        p = re.sub(r"^[A-Da-d][.)]\s*", "", p)
        if p:
            out.append(p)
    return out


def extract_labeled_value(text: str, label_patterns: Iterable[str]) -> Optional[str]:
    """Pull 'Label: value' / 'Label - value' from a blended cell."""
    raw = norm(text)
    if not raw:
        return None
    # Longer labels first so 'Account Name' wins over 'Account'
    labels = sorted(label_patterns, key=lambda x: len(x), reverse=True)
    for label in labels:
        # Prefer label at start of cell — avoid mid-paragraph hits like "...rebate incentive..."
        for pat in (
            rf"(?i)^\s*{re.escape(label)}\s*[:\-–—]\s*(.+)$",
            rf"(?i)^\s*{re.escape(label)}\s+(.+)$",
            rf"(?i)\b{re.escape(label)}\s*[:\-–—]\s*(.+)$",
        ):
            m = re.search(pat, raw)
            if m:
                # Mid-body matches on long narrative cells are usually false positives
                if m.start() > 0 and len(raw) > 120:
                    continue
                val = norm(m.group(1))
                # Reject captures that are just the rest of a longer label ("Name:")
                if re.fullmatch(r"[A-Za-z][A-Za-z ]*:?", val) and len(val) <= 20:
                    # likely label fragment unless it looks like a real short value
                    if ":" in val or val.lower() in {
                        "name",
                        "name:",
                        "manager",
                        "status",
                        "rank",
                    }:
                        continue
                return val
    return None


def looks_like_label_only(text: str) -> bool:
    s = norm(text)
    if not s or len(s) > 80:
        return False
    if re.search(r"[:：]\s*\S+", s):
        return False
    return bool(re.match(r"^[A-Za-z0-9][A-Za-z0-9 &'/%()+.,\-]{0,70}$", s))


# ---------------------------------------------------------------------------
# SFXL (Salesforce report)
# ---------------------------------------------------------------------------

@dataclass
class SfxlRow:
    raw: dict[str, str]


def _find_header_row(rows: list[list[Any]], required_hints: list[str]) -> Optional[int]:
    hints = [norm_key(h) for h in required_hints]
    best_i, best_score = None, 0
    for i, row in enumerate(rows[:40]):
        keys = {norm_key(c) for c in row if c is not None and str(c).strip()}
        score = sum(1 for h in hints if any(h in k or k in h for k in keys))
        if score > best_score:
            best_score, best_i = score, i
    return best_i if best_score >= 2 else None


def _map_columns(headers: list[str]) -> dict[str, int]:
    """Map logical fields to header indexes. Exact alias match wins over contains."""
    mapped: dict[str, int] = {}
    header_norms = [(i, norm_key(h)) for i, h in enumerate(headers)]
    for logical, aliases in SFXL_COLUMN_ALIASES.items():
        alias_norms = [norm_key(a) for a in aliases]
        # 1) exact header == alias
        for i, hk in header_norms:
            if hk in alias_norms:
                mapped[logical] = i
                break
        if logical in mapped:
            continue
        # 2) header startswith alias + space (rare)
        for a in alias_norms:
            for i, hk in header_norms:
                if hk.startswith(a + " ") or hk.endswith(" " + a):
                    mapped[logical] = i
                    break
            if logical in mapped:
                break
    return mapped


def is_corporate_account_name(account_name: str) -> bool:
    return bool(CORPORATE_NAME_RE.search(norm(account_name)))


def builder_core_name(account_name: str) -> str:
    """
    Normalize builder identity for matching PGS <-> SFXL.
    'Shea Homes - Corporate' / 'DR Horton (DHI)' -> 'shea homes' / 'dr horton'
    """
    s = norm(account_name)
    # Drop parenthetical nicknames: (DHI), (prev. Gehan), etc.
    s = re.sub(r"\([^)]*\)", " ", s)
    s = norm_key(s)
    s = re.sub(r"\bcorp\b", "corporate", s)
    s = re.sub(r"\bcorporate(?:\s+hq)?\b", " ", s)
    s = re.sub(r"\bhq\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" -_")
    return s


def builder_match_keys(account_name: str) -> list[str]:
    """
    Ordered keys for matching a PGS builder name to SFXL Account Names.
    'DR Horton (DHI)' -> ['dr horton', 'dr horton dhi'] (broader first for contains checks)
    """
    keys: list[str] = []
    raw = norm(account_name)
    no_paren = re.sub(r"\([^)]*\)", " ", raw)
    no_paren = re.sub(r"\s+", " ", no_paren).strip()
    for candidate in (no_paren, raw):
        core = builder_core_name(candidate)
        if core and core not in keys:
            keys.append(core)
    return keys


def sfxl_name_matches_builder(sfxl_name: str, builder_name: str) -> bool:
    sfxl_k = norm_key(sfxl_name)
    if not sfxl_k:
        return False
    for key in builder_match_keys(builder_name):
        if key and key in sfxl_k:
            return True
    return False


def load_sfxl(
    path: Path,
    sheet: Optional[str] = None,
    corporate_only: bool = True,
) -> tuple[list[str], list[SfxlRow], dict[str, int]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    if not rows:
        raise ValueError(f"SFXL is empty: {path}")

    header_i = _find_header_row(
        rows,
        ["Account Name", "Website", "Type", "Physical State/Province", "Last Modified Date"],
    )
    if header_i is None:
        header_i = 0

    headers = [cell_str(c) if c is not None else f"COL_{i}" for i, c in enumerate(rows[header_i])]
    colmap = _map_columns(headers)
    if "Account Name" not in colmap:
        raise ValueError(
            "Could not find an 'Account Name' column in SFXL. "
            f"Headers seen: {headers}"
        )

    data: list[SfxlRow] = []
    skipped_non_corporate = 0
    for row in rows[header_i + 1 :]:
        if not any(c is not None and str(c).strip() for c in row):
            continue
        raw = {headers[i]: cell_str(row[i]) if i < len(row) else "" for i in range(len(headers))}
        name = raw.get(headers[colmap["Account Name"]], "")
        if corporate_only and not is_corporate_account_name(name):
            skipped_non_corporate += 1
            continue
        data.append(SfxlRow(raw=raw))

    if corporate_only:
        print(
            f"SFXL: kept {len(data)} Corporate account(s); "
            f"skipped {skipped_non_corporate} non-Corporate row(s)."
        )
    return headers, data, colmap


def sfxl_get(row: SfxlRow, logical: str, colmap: dict[str, int], headers: list[str]) -> str:
    if logical in colmap:
        h = headers[colmap[logical]]
        return row.raw.get(h, "")
    aliases = [norm_key(a) for a in SFXL_COLUMN_ALIASES.get(logical, [logical])]
    for k, v in row.raw.items():
        if norm_key(k) in aliases:
            return v
    return ""


def _match_sfxl_row(
    rows: list[SfxlRow],
    headers: list[str],
    colmap: dict[str, int],
    account_name: str,
) -> Optional[SfxlRow]:
    """
    Match PGS builder/account name within a given SFXL row pool.

    Priority:
      1. Exact Account Name match
      2. Exact core-name match (Shea Homes == Shea Homes - Corporate)
      3. Closest startswith/contains core-name match
    """
    target_raw = norm(account_name)
    target = norm_key(target_raw)
    target_core = builder_core_name(target_raw)
    if not target_core and not target:
        return None

    exact: list[SfxlRow] = []
    core_exact: list[SfxlRow] = []
    fuzzy: list[tuple[int, SfxlRow, str]] = []

    for row in rows:
        name = sfxl_get(row, "Account Name", colmap, headers)
        nk = norm_key(name)
        core = builder_core_name(name)
        if not nk:
            continue
        if nk == target or nk == norm_key(target_raw + " - Corporate"):
            exact.append(row)
            continue
        if target_core and core == target_core:
            core_exact.append(row)
            continue
        if target_core and (core.startswith(target_core) or target_core.startswith(core)):
            fuzzy.append((abs(len(core) - len(target_core)), row, name))

    if exact:
        return exact[0]
    if core_exact:
        # Prefer the shortest HQ-style name if several cores match
        return sorted(
            core_exact,
            key=lambda r: len(sfxl_get(r, "Account Name", colmap, headers)),
        )[0]
    if fuzzy:
        fuzzy.sort(key=lambda x: (x[0], len(x[2])))
        best = fuzzy[0]
        ties = [f for f in fuzzy if f[0] == best[0]]
        if len(ties) > 1:
            names = [sfxl_get(t[1], "Account Name", colmap, headers) for t in ties]
            print(
                f"WARNING: Ambiguous SFXL match for '{account_name}': {names}. "
                "Using the closest name; pass --account to disambiguate."
            )
        return best[1]
    return None


def pick_sfxl_row(
    rows: list[SfxlRow],
    headers: list[str],
    colmap: dict[str, int],
    account_name: str,
    fallback_rows: Optional[list[SfxlRow]] = None,
) -> Optional[SfxlRow]:
    """
    Match PGS builder/account name to the SFXL HQ/Corporate row.

    Tries Corporate-named accounts first. If none match (some builders omit
    "Corporate" in Account Name), retries against the unfiltered report.
    """
    target_core = builder_core_name(account_name)
    if not target_core and not norm_key(account_name):
        print("WARNING: No PGS account name to match against SFXL.")
        return None

    chosen = _match_sfxl_row(rows, headers, colmap, account_name)
    source = "Corporate"
    if chosen is None and fallback_rows:
        print(
            f"No Corporate SFXL account matched PGS '{account_name}' "
            f"(core='{target_core}'). Retrying without the Corp name filter..."
        )
        chosen = _match_sfxl_row(fallback_rows, headers, colmap, account_name)
        source = "all accounts"

    if chosen is None:
        sample_pool = fallback_rows or rows
        available = [
            sfxl_get(r, "Account Name", colmap, headers) for r in sample_pool[:20]
        ]
        print(
            f"ERROR: No SFXL account matched PGS account '{account_name}' "
            f"(core='{target_core}'). Sample names: {available}"
        )
        return None

    matched_name = sfxl_get(chosen, "Account Name", colmap, headers)
    print(f"SFXL match ({source}): PGS '{account_name}' -> '{matched_name}'")
    return chosen

# ---------------------------------------------------------------------------
# PGS (prior gold sheet) — label/value layout with merged / blended cells
# ---------------------------------------------------------------------------

@dataclass
class PgsData:
    label_map: dict[str, str] = field(default_factory=dict)
    cells: list[tuple[int, int, str]] = field(default_factory=list)  # row, col, text
    sheet_name: str = ""


def load_pgs(path: Path, sheet: Optional[str] = None) -> PgsData:
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]

    # Expand merged cells into a virtual grid of values
    merged_values: dict[tuple[int, int], Any] = {}
    for mr in ws.merged_cells.ranges:
        min_r, min_c, max_r, max_c = mr.min_row, mr.min_col, mr.max_row, mr.max_col
        top_left = ws.cell(min_r, min_c).value
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                merged_values[(r, c)] = top_left

    cells: list[tuple[int, int, str]] = []
    grid: dict[tuple[int, int], str] = {}
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = merged_values.get((r, c), ws.cell(r, c).value)
            text = cell_str(val)
            if text:
                cells.append((r, c, text))
                grid[(r, c)] = text

    label_map: dict[str, str] = {}

    # Pass 1: classic label | value (right / below)
    for r, c, text in cells:
        label_part, inline_val = _split_inline_kv(text)
        key = norm_key(label_part if label_part else text)

        if inline_val:
            label_map.setdefault(key, inline_val)
            # also store under full original label key
            label_map.setdefault(norm_key(text.split(":")[0]), inline_val)
            continue

        if not looks_like_label_only(text) and ":" not in text:
            continue

        # Right neighbor
        right = _first_nonempty_right(grid, r, c, max_col)
        below = _first_nonempty_below(grid, r, c, max_row, max_scan=8)
        if right and norm_key(right) != key:
            label_map.setdefault(key, right)
        elif below and norm_key(below) != key:
            label_map.setdefault(key, below)

    # Pass 2: keep raw cell texts under their own keys for block extraction
    for r, c, text in cells:
        label_map.setdefault(f"__cell_{r}_{c}", text)

    wb.close()
    return PgsData(label_map=label_map, cells=cells, sheet_name=ws.title)


def _split_inline_kv(text: str) -> tuple[str, str]:
    """
    'Rebate Structure: 3% ...' -> ('Rebate Structure', '3% ...')
    'Total number of Divisions: 12' -> ('Total number of Divisions', '12')
    """
    m = re.match(r"^(.{2,80}?)\s*[:：]\s*(.+)$", norm(text), flags=re.DOTALL)
    if m:
        return norm(m.group(1)), norm(m.group(2))
    return norm(text), ""


def _first_nonempty_right(grid: dict[tuple[int, int], str], r: int, c: int, max_col: int) -> str:
    for cc in range(c + 1, min(c + 6, max_col + 1)):
        if (r, cc) in grid:
            return grid[(r, cc)]
    return ""


def _first_nonempty_below(
    grid: dict[tuple[int, int], str], r: int, c: int, max_row: int, max_scan: int = 8
) -> str:
    for rr in range(r + 1, min(r + max_scan + 1, max_row + 1)):
        if (rr, c) in grid:
            return grid[(rr, c)]
        # sometimes value is one column to the right on next rows
        if (rr, c + 1) in grid:
            return grid[(rr, c + 1)]
    return ""


def _label_matches_alias(alias_key: str, label_key: str) -> bool:
    """Match PGS labels to aliases without false hits (e.g. NAM inside 'account name')."""
    if not alias_key or not label_key:
        return False
    if alias_key == label_key:
        return True
    # Never treat numeric / money cells as labels (e.g. "20" vs "2026 Pricing...")
    if re.fullmatch(r"[\d.,$%+\-/#]+", label_key):
        return False
    # Long narrative cells are values, not labels
    if len(label_key) > 60:
        return False
    if len(alias_key) <= 8:
        return bool(
            re.search(rf"(?:^|\s){re.escape(alias_key)}(?:\s|$)", label_key)
            and len(label_key) <= max(40, len(alias_key) + 20)
        )
    # Longer aliases: require a meaningful prefix, not a 1–2 digit coincidence
    if label_key.startswith(alias_key):
        return True
    if alias_key.startswith(label_key) and len(label_key) >= 8:
        return True
    return False


def _value_looks_incomplete(value: str) -> bool:
    """True when a cell looks like a partial value (e.g. '#' alone) needing the next cell."""
    s = norm(value)
    if not s:
        return True
    if s in {"#", "-", "—", "–", ":", "N/A", "n/a"}:
        return True
    if re.fullmatch(r"#\s*", s):
        return True
    return False


def _is_pgs_label_like(text: str) -> bool:
    label_part, inline = _split_inline_kv(text)
    if inline:
        return True
    raw = norm(text)
    if raw.endswith(":") or raw.endswith("："):
        return True
    if looks_like_label_only(label_part) and any(
        token in norm_key(label_part)
        for token in (
            "strategy",
            "positioning",
            "selection point",
            "description",
            "brands",
            "program",
            "pricing",
            "terms",
            "divisions",
            "account manager",
        )
    ):
        return True
    nk = norm_key(label_part)
    for aliases in LABEL_ALIASES.values():
        if any(_label_matches_alias(norm_key(a), nk) for a in aliases):
            return True
    return False


def _cells_right_of(pgs: PgsData, row: int, col: int) -> list[str]:
    return [
        norm(t)
        for r, c, t in sorted(pgs.cells, key=lambda x: x[1])
        if r == row and c > col and norm(t)
    ]


def pgs_lookup(pgs: PgsData, logical: str) -> str:
    """
    Resolve a PGS field value.
    Prefer the label row's next cells (handles 'Builder Rank #' | '1').
    If a candidate value is incomplete, keep reading the next cell.
    """
    aliases = LABEL_ALIASES.get(logical, [logical])
    alias_keys = [norm_key(a) for a in aliases]

    if logical in BLOCK_VALUE_LOGICALS and logical != "Builder Strategy":
        block_values = extract_label_block_values(pgs, logical)
        if block_values:
            return " ".join(block_values).strip()

    anchor = find_label_anchor(pgs, logical)
    if anchor:
        r0, c0, label_text = anchor
        parts: list[str] = []

        # Value embedded in the label cell after the label text, e.g. "Builder Rank # 13"
        for alias in aliases:
            embedded = extract_labeled_value(label_text, [alias])
            if embedded and not _value_looks_incomplete(embedded):
                # Avoid treating a lone trailing "#" as the full value
                if not (embedded.strip() == "#" or re.fullmatch(r"#\s*", embedded)):
                    return embedded
            if embedded and _value_looks_incomplete(embedded):
                parts.append(embedded)
                break
            # Label cell ends with "#" after alias words: "Builder Rank #"
            m = re.search(
                rf"(?i)\b{re.escape(alias)}\s*(#)\s*$",
                norm(label_text),
            )
            if m:
                parts.append("#")
                break

        for cell in _cells_right_of(pgs, r0, c0):
            if _is_pgs_label_like(cell) and not parts:
                continue
            if norm_key(cell) in alias_keys:
                continue
            # Skip repeating the label text itself ("Builder Rank #")
            if any(_label_matches_alias(ak, norm_key(cell)) for ak in alias_keys):
                # but keep a trailing "#" fragment
                if cell.strip().endswith("#") and not re.search(r"\d", cell):
                    if "#" not in parts:
                        parts.append("#")
                    continue
                continue
            parts.append(cell)
            joined = " ".join(parts).strip()
            if not _value_looks_incomplete(joined):
                return _cleanup_pgs_value(joined, logical)

        if parts:
            return _cleanup_pgs_value(" ".join(parts), logical)
        if logical in STRICT_ADJACENT_VALUE_LOGICALS:
            return ""

    # Fallback: label_map / inline extraction, then next-cell repair via anchor if needed
    for ak in alias_keys:
        for k, v in pgs.label_map.items():
            if k.startswith("__cell_"):
                continue
            if _label_matches_alias(ak, k):
                if v and norm_key(v) not in alias_keys:
                    if _value_looks_incomplete(v) and anchor:
                        nxt = _cells_right_of(pgs, anchor[0], anchor[1])
                        if nxt:
                            return _cleanup_pgs_value(f"{v} {nxt[0]}", logical)
                    return _cleanup_pgs_value(v, logical)

    for _, _, text in pgs.cells:
        extracted = extract_labeled_value(text, aliases)
        if extracted and not _value_looks_incomplete(extracted):
            return _cleanup_pgs_value(extracted, logical)

    return ""


def _cleanup_pgs_value(value: str, logical: str) -> str:
    s = re.sub(r"\s+", " ", norm(value)).strip()
    if logical == "Builder Rank":
        m = re.search(r"#\s*(\d+)", s)
        if m:
            return f"# {m.group(1)}"
        m = re.fullmatch(r"(\d+)", s)
        if m:
            return f"# {m.group(1)}"
    return s


def find_label_anchor(pgs: PgsData, logical: str) -> Optional[tuple[int, int, str]]:
    """
    Prefer short label cells (e.g. 'Rebate Structure:') over long paragraphs that
    merely mention a word like 'rebate' mid-sentence.
    """
    aliases = [norm_key(a) for a in LABEL_ALIASES.get(logical, [logical])]
    best: Optional[tuple[int, int, int, str]] = None  # score, row, col, text
    for r, c, text in pgs.cells:
        label_part, inline = _split_inline_kv(text)
        nk = norm_key(label_part)
        if not any(_label_matches_alias(a, nk) for a in aliases):
            continue
        # Score: exact key > short label-like cell > anything else
        score = 0
        if any(nk == a for a in aliases):
            score += 100
        if looks_like_label_only(label_part) or (not inline and norm(text).endswith((":", "："))):
            score += 50
        if len(norm(text)) <= 80:
            score += 25
        if inline:
            score += 10
        # Heavy penalty for narrative paragraphs
        if len(norm(text)) > 120:
            score -= 80
        if best is None or score > best[0]:
            best = (score, r, c, text)
    if best and best[0] >= 50:
        return best[1], best[2], best[3]
    return None


def extract_builder_strategies(pgs: PgsData, max_items: int = 4) -> list[str]:
    """
    Builder's Strategy is often one merged/multi-line cell, or a label with
    4 numbered items in following rows/cells.
    """
    row_values = extract_label_block_values(pgs, "Builder Strategy")
    if row_values:
        seen: set[str] = set()
        out: list[str] = []
        for value in row_values:
            for piece in split_multiline(value):
                k = norm_key(piece)
                if not k or k in seen:
                    continue
                seen.add(k)
                out.append(piece)
                if len(out) >= max_items:
                    return out
        if out:
            return out

    anchor = find_label_anchor(pgs, "Builder Strategy")
    collected: list[str] = []

    if anchor:
        r0, c0, text = anchor
        _, inline = _split_inline_kv(text)
        if inline:
            collected.extend(split_multiline(inline))
        # If the cell itself is only the label, gather following content nearby
        block_texts: list[str] = []
        for r, c, t in pgs.cells:
            if r < r0:
                continue
            if r == r0 and c < c0:
                continue
            if r > r0 + 12:
                break
            if abs(c - c0) <= 3:
                label_part, _ = _split_inline_kv(t)
                label_key = norm_key(label_part)
                # skip the Builder Strategy header cell itself
                if label_key in {norm_key(a) for a in LABEL_ALIASES["Builder Strategy"]}:
                    continue
                # Once collection has started, stop before crossing into the next labeled section.
                if block_texts and _is_pgs_label_like(label_part):
                    break
                block_texts.append(t)
        for t in block_texts:
            # stop if we hit another major section label
            if _is_other_section_header(t, exclude="Builder Strategy"):
                break
            collected.extend(split_multiline(t))

    # Fallback: any cell containing "strategy" with multiple lines
    if len(collected) < 2:
        for _, _, t in pgs.cells:
            if "strateg" in norm_key(t) and ("\n" in t or re.search(r"\n|^\s*1[.)]", t)):
                _, inline = _split_inline_kv(t)
                collected = split_multiline(inline or t)
                if collected:
                    # drop the header word if present
                    collected = [x for x in collected if "strateg" not in norm_key(x)]
                    break

    # Deduplicate while preserving order
    seen = set()
    uniq: list[str] = []
    for item in collected:
        k = norm_key(item)
        if not k or k in seen:
            continue
        if k in {norm_key(a) for a in LABEL_ALIASES["Builder Strategy"]}:
            continue
        seen.add(k)
        uniq.append(item)

    return uniq[:max_items]


def extract_label_block_values(pgs: PgsData, logical: str) -> list[str]:
    """
    Return the row-by-row values to the right of a vertically merged label block.

    Example:
      A14:A16 = Builder's Strategy
      B14:D14 = point 1
      B15:D15 = point 2
      B16:D16 = point 3
    -> [point 1, point 2, point 3]
    """
    anchor = find_label_anchor(pgs, logical)
    if not anchor:
        return []

    r0, c0, label_text = anchor
    label_key = norm_key(_split_inline_kv(label_text)[0])

    block_rows: list[int] = []
    for r, c, t in sorted(pgs.cells, key=lambda x: (x[0], x[1])):
        if c != c0 or r < r0:
            continue
        cell_key = norm_key(_split_inline_kv(t)[0])
        if cell_key == label_key or norm_key(t) == label_key:
            block_rows.append(r)
            continue
        if block_rows:
            break

    if not block_rows:
        block_rows = [r0]

    values: list[str] = []
    seen: set[str] = set()
    for r in block_rows:
        right = sorted(
            ((c, t) for rr, c, t in pgs.cells if rr == r and c > c0),
            key=lambda x: x[0],
        )
        if not right:
            continue
        val = norm(right[0][1])
        key = norm_key(val)
        if not key or key == label_key or key in seen:
            continue
        if _is_other_section_header(val, exclude=logical):
            break
        seen.add(key)
        values.append(val)
    return values


def parse_operating_divisions_text(text: str) -> tuple[list[str], str]:
    """
    Parse a PGS Operating Divisions value such as:
      '15 Divisions, 9 States: AZ, CO, CA, FL, NV, NC, TX, VA, WA'
    -> items=['AZ','CO',...], total='15'
    """
    raw = norm(text)
    if not raw:
        return [], ""

    total = ""
    m_total = re.search(r"(?i)\b(\d{1,4})\s+divisions?\b", raw)
    if m_total:
        total = m_total.group(1)
    else:
        m_total = re.search(
            r"(?i)\b(?:total(?:\s+number\s+of)?\s+divisions?|total\s*d)\s*[:\-–—]?\s*(\d{1,4})\b",
            raw,
        )
        if m_total:
            total = m_total.group(1)

    items: list[str] = []
    m_states = re.search(r"(?i)\bstates?\s*[:\-–—]\s*(.+)$", raw)
    if m_states:
        list_blob = m_states.group(1)
    else:
        # Strip leading summary clause like "15 Divisions, 9 States:" if present without match above
        list_blob = re.sub(
            r"(?i)^\s*\d{1,4}\s+divisions?\s*(?:,\s*\d{1,4}\s+states?\s*)?[:\-–—]?\s*",
            "",
            raw,
        ).strip()
        if list_blob == raw and total:
            # value was only "15" / "15 Divisions"
            list_blob = ""

    if list_blob:
        for piece in re.split(r"[,;/|]+|\n+|•|\u2022", list_blob):
            piece = norm(piece)
            piece = re.sub(r"^\d+[.)]\s*", "", piece)
            if not piece:
                continue
            # skip leftover summary fragments
            if re.fullmatch(r"(?i)\d{1,4}\s+divisions?", piece):
                continue
            if re.fullmatch(r"(?i)\d{1,4}\s+states?", piece):
                continue
            if "builder brand" in norm_key(piece) or norm_key(piece) == "description":
                continue
            items.append(piece)

    return items, total


def extract_label_block_joined(
    pgs: PgsData,
    logical: str,
    joiner: str = "\n",
) -> str:
    """
    Collect ALL value cells under a vertically merged PGS label into one string.

    Example — label A41:A43 = '2026 Pricing Action Status' with:
      B41 Price increase...
      B42 Installation Included
      B43 Facings Excluded
    -> joined text (GrowthStrat* stays split via extract_builder_strategies instead).
    """
    values = extract_label_block_values(pgs, logical)
    if values:
        return joiner.join(values)
    return pgs_lookup(pgs, logical)


def sfxl_last_modified_date(
    row: Optional[SfxlRow],
    colmap: dict[str, int],
    headers: list[str],
) -> str:
    """
    Always pull Date from SFXL column 'Last Modified Date' on the matched account row.
    Never use Last Activity or other date-like columns.
    """
    if not row:
        return ""
    # Prefer exact header key on the row payload.
    for key in ("Last Modified Date", "LastModifiedDate"):
        if key in row.raw and row.raw[key]:
            return row.raw[key]
    if "Last Modified Date" in colmap:
        return sfxl_get(row, "Last Modified Date", colmap, headers)
    print(
        "WARNING: SFXL has no 'Last Modified Date' column; Date will be empty. "
        f"Headers: {headers}"
    )
    return ""


def extract_operating_divisions(pgs: PgsData, max_buckets: int = 4) -> tuple[list[str], str]:
    """
    Returns (division_buckets[4], total_divisions_str).

    Prefers the single value cell beside 'Operating Divisions' (does not bleed into
    Builder Brands / Description / strategy blocks).
    """
    stop_labels = {
        norm_key(x)
        for x in (
            "Builder Brands",
            "Builder Brand",
            "Description",
            "Builder's Strategy",
            "Builders Strategy",
            "Builder Strategy",
            "Product Mix",
            "Home Type",
            "Spec Home %",
            "National Average Selling Price",
            "HHT Sales Strategy/Positioning",
        )
    }

    # 1) Direct lookup of the Operating Divisions value cell
    direct = pgs_lookup(pgs, "Operating Divisions")
    items, total = parse_operating_divisions_text(direct)

    # 2) If lookup returned the label itself or empty, gather same-row / immediate neighbors only
    if not items and not total:
        anchor = find_label_anchor(pgs, "Operating Divisions")
        if anchor:
            r0, c0, _ = anchor
            same_row_bits: list[str] = []
            for r, c, t in pgs.cells:
                if r != r0 or c <= c0:
                    continue
                if norm_key(_split_inline_kv(t)[0]) in stop_labels:
                    break
                same_row_bits.append(t)
            items, total = parse_operating_divisions_text(" ".join(same_row_bits))

    # 3) Optional: short vertical list of division names under the label (same column family),
    #    stopping at the next section header. Only used if still empty.
    if not items:
        anchor = find_label_anchor(pgs, "Operating Divisions")
        if anchor:
            r0, c0, _ = anchor
            collected: list[str] = []
            for r, c, t in sorted(pgs.cells):
                if r <= r0:
                    continue
                if r > r0 + 8:
                    break
                if abs(c - c0) > 2 and abs(c - (c0 + 1)) > 1:
                    continue
                label_part, inline = _split_inline_kv(t)
                if norm_key(label_part) in stop_labels:
                    break
                if _is_other_section_header(t, exclude="Operating Divisions"):
                    break
                blob = inline or t
                # skip narrative paragraphs
                if len(blob) > 120 and "," not in blob:
                    break
                parsed_items, parsed_total = parse_operating_divisions_text(blob)
                if parsed_total and not total:
                    total = parsed_total
                if parsed_items:
                    collected.extend(parsed_items)
                else:
                    for piece in split_multiline(blob):
                        if piece and norm_key(piece) not in stop_labels:
                            collected.append(piece)
            items = collected

    # Dedup preserve order
    seen = set()
    uniq: list[str] = []
    for it in items:
        k = norm_key(it)
        if not k or k in seen:
            continue
        seen.add(k)
        uniq.append(it)

    if not total and uniq:
        # only use count-as-total when values look like named divisions, not USPS states alone
        total = str(len(uniq))

    return _bucket_list(uniq, max_buckets), total


def strip_builder_from_account_name(account_name: str, builder_name: str) -> str:
    """
    'Shea Homes - CA Central Coast' + builder 'Shea Homes' -> 'CA Central Coast'
    """
    name = norm(account_name)
    builder = norm(builder_name)
    if not name:
        return ""
    if builder:
        name = re.sub(re.escape(builder), "", name, count=1, flags=re.IGNORECASE)
    name = re.sub(r"^[\s\-_:/|]+", "", name)
    name = re.sub(r"[\s\-_:/|]+$", "", name)
    return norm(name)


def format_currency_usd(value: str) -> str:
    """Format ASP-like values as $100,000."""
    s = norm(value)
    if not s:
        return ""
    cleaned = s.replace("$", "").replace(",", "").strip()
    try:
        num = float(cleaned)
    except ValueError:
        return s if s.startswith("$") else s
    if num.is_integer():
        return f"${int(num):,}"
    return f"${num:,.2f}"


def format_integer_commas(value: str) -> str:
    """Format whole numbers with thousands separators: 83654 -> 83,654."""
    s = norm(value)
    if not s:
        return ""
    cleaned = s.replace(",", "").strip()
    try:
        num = float(cleaned)
    except ValueError:
        return s
    if num.is_integer():
        return f"{int(num):,}"
    return f"{num:,.2f}"


def build_divisions_from_sfxl(
    rows: list[SfxlRow],
    headers: list[str],
    colmap: dict[str, int],
    builder_name: str,
    max_buckets: int = 4,
) -> tuple[list[str], str, list[str]]:
    """
    Divisions_* / Total_D from SFXL Account Name column:
      1. No Corp-only filter (use full report)
      2. Keep rows whose Account Name matches builder (parentheticals ignored)
      3. Skip Corporate accounts (not 'Corpus')
      4. Strip Builder Name from each remaining Account Name
      5. Print list for review
      6. Split across Divisions_1..4 (total ÷ 4; remainder spread across first columns)
      7. Total_D = count of those entries
    """
    match_keys = builder_match_keys(builder_name)
    strip_name = match_keys[0] if match_keys else norm(builder_name)
    matched_raw: list[str] = []

    for row in rows:
        name = sfxl_get(row, "Account Name", colmap, headers)
        if not name:
            continue
        if not sfxl_name_matches_builder(name, builder_name):
            continue
        if is_corporate_account_name(name):
            continue
        matched_raw.append(name)

    cleaned: list[str] = []
    seen: set[str] = set()
    for name in matched_raw:
        label = strip_builder_from_account_name(name, strip_name)
        # Also try stripping original PGS name if still prefixed oddly
        if not label or norm_key(label) == norm_key(name):
            label = strip_builder_from_account_name(name, builder_name)
        if not label:
            continue
        key = norm_key(label)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(label)

    print("\n=== Divisions candidates (SFXL Account Name, Corp skipped, builder stripped) ===")
    print(f"  Builder match keys: {match_keys}")
    if not cleaned:
        print("  (none)")
    else:
        for i, item in enumerate(cleaned, 1):
            print(f"  {i}. {item}")
    print(f"Total_D (count): {len(cleaned)}")

    buckets = _bucket_list_remainder_first(cleaned, max_buckets)
    return buckets, str(len(cleaned)) if cleaned else "0", cleaned


def _pgs_cell_text(pgs: PgsData, row: int, col: int) -> str:
    for r, c, t in pgs.cells:
        if r == row and c == col:
            return norm(t)
    return ""


def extract_nam_from_pgs(pgs: PgsData) -> tuple[str, str]:
    """
    PGS row: National Account Manager | NAM_Name | NAM_Email | phone
    Name is the cell immediately right of the label; email is the next cell.
    """
    anchor = find_label_anchor(pgs, "National Account Manager")
    if not anchor:
        return "", ""

    r0, c0, _ = anchor
    name = _pgs_cell_text(pgs, r0, c0 + 1)
    email = _pgs_cell_text(pgs, r0, c0 + 2)

    # If email column is phone, scan further right for @
    if email and "@" not in email:
        for cc in range(c0 + 2, c0 + 6):
            candidate = _pgs_cell_text(pgs, r0, cc)
            if "@" in candidate:
                email = candidate
                break

    # Guard: name column must not be the builder account name from a bad row match
    if name and "@" in name:
        if not email:
            email = name
        name = ""

    return name, email


def _division_column_sizes(count: int, columns: int = 4) -> list[int]:
    """Divide count across columns; +1 item on the first `remainder` columns."""
    if count <= 0:
        return [0] * columns
    base, rem = divmod(count, columns)
    return [base + (1 if i < rem else 0) for i in range(columns)]


def format_division_bucket(entries: list[str], start_number: int = 1) -> str:
    """Legacy helper — prefer _bucket_list_remainder_first for Divisions_*."""
    cleaned = [norm(e) for e in entries if norm(e)]
    if not cleaned:
        return ""
    numbered = [f"{start_number + i}. {entry}" for i, entry in enumerate(cleaned)]
    return "\n" + "\n\n".join(numbered)


def _bucket_list_remainder_first(items: list[str], n: int = 4) -> list[str]:
    """
    Split divisions across Divisions_1..4:
      - total ÷ 4 (remainder spread on first columns, e.g. 98 -> 25,25,24,24)
      - number ALL items 1..total globally, then split (continuous numbering)
      - tighter line spacing when many entries so columns can show more rows
    """
    if n <= 0:
        return []
    cleaned = [norm(x) for x in items if norm(x)]
    if not cleaned:
        return [""] * n

    total = len(cleaned)
    sizes = _division_column_sizes(total, n)
    numbered = [f"{i}. {text}" for i, text in enumerate(cleaned, start=1)]

    tight = total > 20
    sep = "\n" if tight else "\n\n"  # blank line between items when list is short enough
    lead = "\n"

    out: list[str] = []
    idx = 0
    for col_i, size in enumerate(sizes, start=1):
        chunk = numbered[idx : idx + size]
        idx += size
        start_num = idx - len(chunk) + 1 if chunk else 0
        end_num = idx if chunk else 0
        print(f"  Divisions_{col_i}: #{start_num}-{end_num} ({len(chunk)} items)")
        out.append(lead + sep.join(chunk) if chunk else "")
    return out


def _bucket_list(items: list[str], n: int) -> list[str]:
    """Split a long list into n roughly equal newline-joined buckets (PDF has 4 division fields)."""
    if not items:
        return [""] * n
    if len(items) <= n:
        return items + [""] * (n - len(items))
    # Contiguous chunks so related regions stay together
    size = (len(items) + n - 1) // n
    out: list[str] = []
    for i in range(n):
        chunk = items[i * size : (i + 1) * size]
        out.append("\n".join(chunk))
    return out


def _is_other_section_header(text: str, exclude: str) -> bool:
    nk = norm_key(_split_inline_kv(text)[0])
    if not looks_like_label_only(_split_inline_kv(text)[0]) and ":" not in text:
        return False
    for logical, aliases in LABEL_ALIASES.items():
        if logical == exclude:
            continue
        for a in aliases:
            if nk == norm_key(a):
                return True
    # common gold-sheet section titles
    headers = {
        "product mix",
        "competitive landscape",
        "notes",
        "comments",
        "hht",
        "hht sales strategy positioning",
        "supplier status",
        "strategic tier",
        "homebuyer selection point",
        "builder selection point for spec homes",
        "builder selection point for model homes",
    }
    return nk in headers


# ---------------------------------------------------------------------------
# Mapping assembly
# ---------------------------------------------------------------------------

def build_field_values(
    pgs: PgsData,
    sfxl_row: Optional[SfxlRow],
    sfxl_headers: list[str],
    sfxl_colmap: dict[str, int],
    sfxl_all_rows: Optional[list[SfxlRow]] = None,
) -> dict[str, str]:
    def sf(logical: str) -> str:
        if not sfxl_row:
            return ""
        return sfxl_get(sfxl_row, logical, sfxl_colmap, sfxl_headers)

    strategies = [
        normalize_display_casing(s) for s in extract_builder_strategies(pgs, max_items=4)
    ]
    while len(strategies) < 4:
        strategies.append("")

    builder_name = pgs_lookup(pgs, "Account Name")
    division_rows = sfxl_all_rows if sfxl_all_rows is not None else (
        [sfxl_row] if sfxl_row else []
    )
    div_buckets, total_d, _ = build_divisions_from_sfxl(
        division_rows, sfxl_headers, sfxl_colmap, builder_name, max_buckets=4
    )
    nam_name, nam_email = extract_nam_from_pgs(pgs)

    values = {
        "Builder_Name": builder_name,
        "Headquarters": sf("Physical State/Province"),
        "Website_URL": sf("Website"),
        "Market_Facing_Position": sf("Type"),
        "Date": date.today().strftime("%m/%d/%Y"),
        "NAM_Name": nam_name,
        "NAM_Email": nam_email,
        "Builder_Rank": pgs_lookup(pgs, "Builder Rank"),
        "Annual_Closings": format_integer_commas(pgs_lookup(pgs, "Annual Closings")),
        "National_ASP": format_currency_usd(
            pgs_lookup(pgs, "National Average Selling Price")
        ),
        "Spec_Percent": format_spec_percent(pgs_lookup(pgs, "Spec Home %")),
        "Rebate_Structure": normalize_display_casing(
            extract_label_block_joined(pgs, "Rebate Structure")
        ),
        "Model_Home_Design_Center_Program": normalize_display_casing(
            extract_label_block_joined(pgs, "Design Center Program")
            or pgs_lookup(pgs, "Design Center Program")
        ),
        "Pricing_Actions": normalize_display_casing(
            extract_label_block_joined(pgs, "2026 Pricing Action Status")
        ),
        "National_Pricing": normalize_display_casing(pgs_lookup(pgs, "National Pricing")),
        "Special_Terms": normalize_display_casing(pgs_lookup(pgs, "Special Terms")),
        "GrowthStrat1_Text": strategies[0],
        "GrowthStrat2_Text": strategies[1],
        "GrowthStrat3_Text": strategies[2],
        "GrowthStrat4_Text": strategies[3],
        # Separate label fields beside each GrowthStrat*_Text box
        "Growth_Strat_1": "Growth Strat 1" if strategies[0] else "",
        "Growth_Strat_2": "Growth Strat 2" if strategies[1] else "",
        "Growth_Strat_3": "Growth Strat 3" if strategies[2] else "",
        "Growth_Strat_4": "Growth Strat 4" if strategies[3] else "",
        "Divisions_1": div_buckets[0],
        "Divisions_2": div_buckets[1],
        "Divisions_3": div_buckets[2],
        "Divisions_4": div_buckets[3],
        "Total_D": total_d,
    }
    return values


# ---------------------------------------------------------------------------
# PDF fill
# ---------------------------------------------------------------------------

def list_pdf_fields(pdf_path: Path) -> dict[str, str]:
    reader = PdfReader(str(pdf_path))
    fields = reader.get_fields() or {}
    out: dict[str, str] = {}
    for name, meta in fields.items():
        if name is None:
            continue
        val = ""
        if isinstance(meta, dict):
            v = meta.get("/V")
            val = "" if v is None else str(v)
        out[str(name)] = val
    return out


def resolve_pdf_field_names(available: Iterable[str]) -> dict[str, str]:
    """logical -> actual PDF field name present in the template."""
    avail = list(available)
    avail_l = {a.lower(): a for a in avail}
    resolved: dict[str, str] = {}
    for logical, candidates in PDF_FIELD_CANDIDATES.items():
        for cand in candidates:
            if cand in avail:
                resolved[logical] = cand
                break
            if cand.lower() in avail_l:
                resolved[logical] = avail_l[cand.lower()]
                break
        if logical not in resolved:
            # soft contains match
            for a in avail:
                if logical.lower() in a.lower() or a.lower() in logical.lower():
                    resolved[logical] = a
                    break
    return resolved


def _set_need_appearances(writer: PdfWriter) -> None:
    try:
        if writer._root_object.get("/AcroForm") is None:  # noqa: SLF001
            writer._root_object.update(  # noqa: SLF001
                {NameObject("/AcroForm"): writer._create_object({})}  # noqa: SLF001
            )
        writer._root_object["/AcroForm"].update(  # noqa: SLF001
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )
    except Exception:
        pass


def format_spec_percent(value: str) -> str:
    """
    Preserve sheet-style percent text (>30%, 70%, etc.).
    Excel often stores 70% as 0.7 with number format 0% — convert those to '70%'.
    """
    s = norm(value)
    if not s:
        return ""
    if "%" in s:
        return s
    cleaned = s.replace(",", "").strip()
    try:
        num = float(cleaned)
    except ValueError:
        return s
    # Decimal percent storage (0.7 -> 70%)
    if 0 <= num <= 1:
        pct = num * 100
        if pct.is_integer():
            return f"{int(pct)}%"
        return f"{pct:g}%"
    # Whole number without symbol (20 -> 20%, 30 -> 30%)
    if num.is_integer():
        return f"{int(num)}%"
    return f"{num:g}%"


_KEEP_ACRONYMS = frozenset(
    {
        "HHT",
        "ASP",
        "NAM",
        "HQ",
        "SFD",
        "DV",
        "BDM",
        "VP",
        "USA",
        "US",
        "NC",
        "VA",
        "GA",
        "TX",
        "FL",
        "AZ",
        "CA",
        "CO",
        "NV",
        "WA",
        "OR",
        "NY",
        "NJ",
        "PA",
        "OH",
        "IN",
        "IL",
        "MI",
        "TN",
        "SC",
        "MD",
        "DE",
        "CT",
        "MA",
        "RI",
        "NH",
        "VT",
        "ME",
        "WI",
        "MN",
        "IA",
        "MO",
        "AR",
        "LA",
        "MS",
        "AL",
        "KY",
        "WV",
        "OK",
        "KS",
        "NE",
        "SD",
        "ND",
        "MT",
        "ID",
        "WY",
        "UT",
        "NM",
        "AK",
        "HI",
        "DC",
    }
)


def normalize_display_casing(text: str) -> str:
    """
    Soften shouty ALL-CAPS runs for PDF readability while keeping short acronyms.
    'PROVIDER OF QUALITY PRODUCTS AND SERVICES delivers...' ->
    'Provider Of Quality Products And Services delivers...'
    """
    s = norm(text)
    if not s:
        return ""

    small_words = frozenset({"OF", "AND", "THE", "A", "AN", "OR", "TO", "IN", "ON", "FOR", "BY"})

    def _fix_word(word: str, *, in_phrase: bool) -> str:
        letters = "".join(ch for ch in word if ch.isalpha())
        if not letters:
            return word
        upper = letters.upper()
        if upper in _KEEP_ACRONYMS:
            return re.sub(r"[A-Za-z]+", upper, word, count=1)
        if letters.isupper() and (len(letters) > 4 or (in_phrase and upper in small_words) or (in_phrase and len(letters) >= 2)):
            titled = letters.title()
            return re.sub(r"[A-Za-z]+", titled, word, count=1)
        return word

    # Multi-word ALL CAPS phrase (2+ consecutive uppercase words)
    def _repl_phrase(m: re.Match[str]) -> str:
        return " ".join(_fix_word(w, in_phrase=True) for w in m.group(0).split())

    s = re.sub(r"\b[A-Z]{2,}(?:\s+[A-Z]{2,})+\b", _repl_phrase, s)
    # Lone long ALL CAPS token
    s = re.sub(
        r"\b[A-Z]{5,}\b",
        lambda m: m.group(0) if m.group(0) in _KEEP_ACRONYMS else m.group(0).title(),
        s,
    )
    return s


def fit_single_line_font_size(
    text: str,
    width: float,
    height: float,
    min_size: float = 10.0,
    max_size: float = 20.0,
) -> float:
    """Largest single-line font that fits the field box height (and width)."""
    s = norm(text)
    if not s:
        return UNIFORM_FIELD_SIZE_LARGE
    size = min(max_size, height * 0.82)
    chars = len(s)
    if chars > 0:
        width_cap = (width * 0.94) / (chars * 0.48)
        size = min(size, width_cap)
    return round(max(min_size, size), 1)


def make_field_da(
    font_size: float,
    *,
    white: bool = False,
    bold: bool = False,
) -> str:
    font = BOLD_FIELD_FONT if bold else UNIFORM_FIELD_FONT
    color = "1 1 1 rg" if white else "0 g"
    return f"/{font} {font_size:g} Tf {color}"


def _estimate_wrapped_lines(text: str, width: float, font_size: float) -> int:
    """Estimate rendered line count with word wrap inside a PDF field box."""
    usable_width = width * 0.94
    char_w = font_size * 0.48
    if char_w <= 0:
        return 1
    chars_per_line = max(1, int(usable_width / char_w))

    total = 0
    blob = str(text).strip()
    if not blob:
        return 0

    paragraphs = blob.splitlines() if "\n" in blob else [blob]
    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if not paragraph:
            continue
        words = paragraph.split()
        line_len = 0
        lines = 1
        for word in words:
            word_len = len(word)
            if line_len == 0:
                line_len = word_len
            elif line_len + 1 + word_len <= chars_per_line:
                line_len += 1 + word_len
            else:
                lines += 1
                line_len = word_len
        total += lines
    return total or 1


def _text_fits_field(
    text: str,
    width: float,
    height: float,
    font_size: float,
) -> bool:
    """True when wrapped text fits vertically in the field (PDF wraps horizontally)."""
    line_count = _estimate_wrapped_lines(text, width, font_size)
    line_h = font_size * 1.22
    total_h = line_count * line_h + font_size * 0.15
    return total_h <= height * 0.92


def auto_fit_font_size(
    text: str,
    width: float,
    height: float,
    base_size: float = 10.0,
    min_size: float = 6.0,
) -> float:
    """Shrink font only when wrapped content would overflow the field height."""
    if not text or not str(text).strip():
        return base_size
    if _text_fits_field(text, width, height, base_size):
        return base_size
    size = base_size
    while size >= min_size:
        if _text_fits_field(text, width, height, size):
            return round(size, 1)
        size -= 0.5
    return min_size


# Multiline fields that should shrink to fit when content overflows the PDF box.
AUTO_FIT_FIELDS = frozenset(
    {
        "Pricing_Actions",
        "Rebate_Structure",
        "National_Pricing",
        "Special_Terms",
        "Model_Home_Design_Center_Program",
        "GrowthStrat1_Text",
        "GrowthStrat2_Text",
        "GrowthStrat3_Text",
        "GrowthStrat4_Text",
    }
)

# Program Details section — share one font size so long Rebate text doesn't look random.
PROGRAM_DETAIL_FIELDS = frozenset(
    {
        "Rebate_Structure",
        "Model_Home_Design_Center_Program",
        "Pricing_Actions",
        "National_Pricing",
        "Special_Terms",
    }
)

GROWTH_TEXT_FIELDS = frozenset(
    {
        "GrowthStrat1_Text",
        "GrowthStrat2_Text",
        "GrowthStrat3_Text",
        "GrowthStrat4_Text",
    }
)

DIVISION_FIELDS = ("Divisions_1", "Divisions_2", "Divisions_3", "Divisions_4")


def _division_entry_count(text: str) -> int:
    """Count numbered division lines in a Divisions_* value."""
    if not text or not str(text).strip():
        return 0
    lines = [ln.strip() for ln in str(text).splitlines() if ln.strip()]
    numbered = [ln for ln in lines if re.match(r"^\d+\.\s+", ln)]
    if numbered:
        return len(numbered)
    parts = [p.strip() for p in re.split(r"\n\s*\n", str(text).strip()) if p.strip()]
    return len(parts) if parts else len(lines)


def divisions_font_size(
    entry_count: int,
    width: float,
    height: float,
    longest_entry_len: int = 0,
    tight_spacing: bool = False,
) -> float:
    """Font size for Divisions_* based on how many numbered rows fit in the field."""
    if entry_count <= 0:
        return 11.0
    if tight_spacing or entry_count > 20:
        visual_rows = 1 + entry_count * 1.15
    else:
        visual_rows = 1 + entry_count + max(0, entry_count - 1)
    usable_h = max(height * 0.92, 40.0)
    size = usable_h / (visual_rows * 1.25)
    chars = max(longest_entry_len, 12)
    max_from_width = (width * 0.94) / (chars * 0.48)
    if max_from_width > 0:
        size = min(size, max_from_width)
    return round(max(7.5, min(13.0, size)), 1)


def _parse_da(da: Any) -> tuple[str, float]:
    """Return (font_name, size) from a PDF /DA string; default Helv 10."""
    text = str(da or "")
    m = re.search(r"/([A-Za-z0-9]+)\s+([\d.]+)\s+Tf", text)
    if m:
        return m.group(1), float(m.group(2))
    return "Helv", 10.0


def _rect_size(rect: Any) -> tuple[float, float]:
    try:
        x0, y0, x1, y1 = (float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3]))
        return abs(x1 - x0), abs(y1 - y0)
    except Exception:
        return 140.0, 330.0


def _longest_division_line(text: str) -> int:
    lines = [ln.strip() for ln in str(text).splitlines() if ln.strip()]
    return max((len(ln) for ln in lines), default=0)


def _widget_size_for_field(writer: PdfWriter, field_name: str) -> tuple[float, float]:
    for widget in _iter_widgets_named(writer, field_name):
        return _rect_size(widget.get("/Rect"))
    return 140.0, 330.0


def _uniform_divisions_font_size(
    writer: PdfWriter,
    resolved: dict[str, str],
    values: dict[str, str],
) -> Optional[float]:
    """One font size for all Divisions_* columns (based on the fullest column)."""
    max_count = 0
    longest = 0
    width, height = 140.0, 330.0
    for logical in DIVISION_FIELDS:
        actual = resolved.get(logical)
        if not actual:
            continue
        text = values.get(logical, "")
        max_count = max(max_count, _division_entry_count(text))
        longest = max(longest, _longest_division_line(text))
        if max_count:
            width, height = _widget_size_for_field(writer, actual)
    if max_count <= 0:
        return None
    total_divisions = sum(
        _division_entry_count(values.get(k, "")) for k in DIVISION_FIELDS
    )
    return divisions_font_size(
        max_count,
        width,
        height,
        longest,
        tight_spacing=total_divisions > 20,
    )


def da_for_logical(
    logical: str,
    font_size: Optional[float] = None,
) -> str:
    if logical in GROWTH_STRAT_LABELS:
        return make_field_da(font_size or UNIFORM_FIELD_SIZE, white=True, bold=True)
    if logical == "Builder_Name":
        return make_field_da(font_size or UNIFORM_FIELD_SIZE_LARGE, white=True)
    if logical in WHITE_FIELD_DA:
        return WHITE_FIELD_DA[logical]
    if font_size is not None:
        return make_field_da(font_size)
    if logical in LARGE_TEXT_LOGICALS:
        return DA_BLACK_14
    return DA_BLACK_10


def _set_widget_da(widget: Any, da: TextStringObject) -> None:
    widget[NameObject("/DA")] = da
    parent = widget.get("/Parent")
    if parent is not None:
        try:
            parent.get_object()[NameObject("/DA")] = da
        except Exception:
            pass


def apply_field_text_styles(
    writer: PdfWriter,
    resolved: dict[str, str],
    logical_keys: Iterable[str],
    field_values: Optional[dict[str, str]] = None,
) -> None:
    """Per-field DA: white headers, 14pt stats, section-uniform auto-fit, Divisions font."""
    values = field_values or {}
    divisions_font = _uniform_divisions_font_size(writer, resolved, values)

    # One shared size for Program Details so Rebate / Model Home / Pricing look consistent.
    program_font: Optional[float] = None
    program_sizes: list[float] = []
    for logical in PROGRAM_DETAIL_FIELDS:
        actual = resolved.get(logical)
        text = values.get(logical, "")
        if not actual or not str(text).strip():
            continue
        width, height = _widget_size_for_field(writer, actual)
        program_sizes.append(
            auto_fit_font_size(
                text, width, height, base_size=UNIFORM_FIELD_SIZE, min_size=8.0
            )
        )
    if program_sizes:
        program_font = min(program_sizes)

    # Growth strategy texts share one size as well.
    growth_font: Optional[float] = None
    growth_sizes: list[float] = []
    for logical in GROWTH_TEXT_FIELDS:
        actual = resolved.get(logical)
        text = values.get(logical, "")
        if not actual or not str(text).strip():
            continue
        width, height = _widget_size_for_field(writer, actual)
        growth_sizes.append(
            auto_fit_font_size(
                text, width, height, base_size=UNIFORM_FIELD_SIZE, min_size=8.0
            )
        )
    if growth_sizes:
        growth_font = min(growth_sizes)

    applied: dict[str, str] = {}
    builder_name_font: Optional[float] = None
    for logical in logical_keys:
        actual = resolved.get(logical)
        if not actual:
            continue

        font_size: Optional[float] = None
        if logical == "Builder_Name":
            text = values.get(logical, "")
            width, height = _widget_size_for_field(writer, actual)
            font_size = fit_single_line_font_size(text, width, height)
            builder_name_font = font_size
        elif logical in DIVISION_FIELDS and divisions_font is not None:
            font_size = divisions_font
        elif logical in PROGRAM_DETAIL_FIELDS and program_font is not None:
            font_size = program_font
        elif logical in GROWTH_TEXT_FIELDS and growth_font is not None:
            font_size = growth_font
        elif logical in AUTO_FIT_FIELDS:
            text = values.get(logical, "")
            width, height = _widget_size_for_field(writer, actual)
            base = UNIFORM_FIELD_SIZE_STAT if logical in LARGE_TEXT_LOGICALS else UNIFORM_FIELD_SIZE
            font_size = auto_fit_font_size(
                text, width, height, base_size=base, min_size=8.0
            )
        elif logical in LARGE_TEXT_LOGICALS:
            font_size = UNIFORM_FIELD_SIZE_STAT
        elif logical in GROWTH_STRAT_LABELS:
            font_size = UNIFORM_FIELD_SIZE

        da_str = da_for_logical(logical, font_size)
        if actual in applied and applied[actual] != da_str:
            continue
        applied[actual] = da_str
        da = TextStringObject(da_str)
        for widget in _iter_widgets_named(writer, actual):
            _set_widget_da(widget, da)

    if divisions_font is not None:
        print(f"  Divisions uniform font: {divisions_font}pt")
    if builder_name_font is not None:
        print(f"  Builder name font: {builder_name_font}pt")
    if program_font is not None:
        print(f"  Program Details uniform font: {program_font}pt")
    if growth_font is not None:
        print(f"  Growth Strategy uniform font: {growth_font}pt")
    print(
        f"Applied field styles: {len(applied)} PDF field(s) "
        f"(white header/labels: {len(WHITE_FIELD_DA) + len(GROWTH_STRAT_LABELS) + 1}, "
        f"stat fields: {UNIFORM_FIELD_SIZE_STAT}pt, "
        f"auto-fit fields: {len(AUTO_FIT_FIELDS)})"
    )


def default_output_path(pgs_path: Path) -> Path:
    """Same base name as the PGS file, e.g. GS - Shea Homes 2026.xlsm -> GS - Shea Homes 2026.pdf"""
    stem = pgs_path.stem.strip()
    stem = re.sub(r"\s*-\s*filled\s*$", "", stem, flags=re.IGNORECASE)
    return pgs_path.with_name(f"{stem}.pdf")


def resolve_output_path(pgs_path: Path, out_override: Optional[Path] = None) -> Path:
    """
    Default: PGS stem + .pdf (no '- filled' suffix).
    If --out is passed with legacy '- filled' in the name, ignore it and use PGS name.
    """
    default = default_output_path(pgs_path)
    if out_override is None:
        return default
    if re.search(r"\s*-\s*filled\s*$", out_override.stem, flags=re.IGNORECASE):
        return default
    return out_override


def _iter_widgets_named(writer: PdfWriter, field_name: str):
    """Yield widget annotation dicts for a field (direct /T or via /Parent)."""
    for page in writer.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        for annot in annots:
            try:
                obj = annot.get_object()
            except Exception:
                continue
            direct = str(obj.get("/T", "") or "")
            parent_name = ""
            parent = obj.get("/Parent")
            if parent is not None:
                try:
                    parent_name = str(parent.get_object().get("/T", "") or "")
                except Exception:
                    parent_name = ""
            if direct == field_name or parent_name == field_name:
                yield obj


def apply_uniform_text_styles(writer: PdfWriter, field_names: list[str]) -> None:
    """Legacy alias — use apply_field_text_styles in fill_pdf instead."""
    for field_name in field_names:
        da = TextStringObject(DA_BLACK_10)
        for widget in _iter_widgets_named(writer, field_name):
            widget[NameObject("/DA")] = da


def fill_pdf(template: Path, field_values: dict[str, str], out_path: Path) -> dict[str, str]:
    reader = PdfReader(str(template))
    writer = PdfWriter()
    writer.append(reader)

    available = list((reader.get_fields() or {}).keys())
    resolved = resolve_pdf_field_names(available)

    pdf_update: dict[str, str] = {}
    unresolved: list[str] = []
    for logical, value in field_values.items():
        actual = resolved.get(logical)
        if not actual:
            unresolved.append(logical)
            continue
        pdf_update[actual] = "" if value is None else str(value)

    if writer.get_fields() is not None:
        for page in writer.pages:
            writer.update_page_form_field_values(page, pdf_update)

    apply_field_text_styles(writer, resolved, field_values.keys(), field_values)

    _set_need_appearances(writer)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "wb") as f:
        writer.write(f)

    if unresolved:
        print("WARNING: PDF fields not found in template:")
        for u in unresolved:
            print(f"  - {u} (tried: {', '.join(PDF_FIELD_CANDIDATES.get(u, []))})")

    return pdf_update


# ---------------------------------------------------------------------------
# Inspect helpers
# ---------------------------------------------------------------------------

def inspect_pgs(path: Path, sheet: Optional[str] = None) -> None:
    pgs = load_pgs(path, sheet)
    print(f"Sheet: {pgs.sheet_name}")
    print(f"Non-empty cells: {len(pgs.cells)}")
    print("\n--- Label map (best-effort) ---")
    for k, v in sorted(pgs.label_map.items()):
        if k.startswith("__cell_"):
            continue
        preview = v if len(v) <= 120 else v[:117] + "..."
        print(f"  [{k}] => {preview}")
    print("\n--- Logical lookups ---")
    for logical in LABEL_ALIASES:
        val = pgs_lookup(pgs, logical)
        preview = val if len(val) <= 120 else val[:117] + "..."
        print(f"  {logical}: {preview or '(empty)'}")
    print("\n--- Builder strategies ---")
    for i, s in enumerate(extract_builder_strategies(pgs), 1):
        print(f"  {i}. {s}")
    divs, total = extract_operating_divisions(pgs)
    print(f"\n--- Operating divisions (Total_D={total}) ---")
    for i, d in enumerate(divs, 1):
        print(f"  Divisions_{i}: {d or '(empty)'}")


def inspect_sfxl(path: Path, sheet: Optional[str] = None, corporate_only: bool = True) -> None:
    headers, rows, colmap = load_sfxl(path, sheet, corporate_only=corporate_only)
    print(f"Rows: {len(rows)}")
    print("Headers:")
    for h in headers:
        print(f"  - {h}")
    print("\nResolved columns:")
    for logical, idx in colmap.items():
        print(f"  {logical} -> {headers[idx]}")
    print("\nCorporate Account Names:")
    for row in rows:
        print(f"  - {sfxl_get(row, 'Account Name', colmap, headers)}")
    if rows:
        print("\nFirst Corporate row sample:")
        for logical in SFXL_COLUMN_ALIASES:
            print(f"  {logical}: {sfxl_get(rows[0], logical, colmap, headers)}")


def inspect_pdf(path: Path) -> None:
    fields = list_pdf_fields(path)
    print(f"Fields: {len(fields)}")
    resolved = resolve_pdf_field_names(fields.keys())
    print("\n--- All PDF field names ---")
    for name in sorted(fields):
        print(f"  {name}")
    print("\n--- Resolved mapping targets ---")
    for logical, actual in resolved.items():
        print(f"  {logical} -> {actual}")
    missing = [k for k in PDF_FIELD_CANDIDATES if k not in resolved]
    if missing:
        print("\n--- Missing ---")
        for m in missing:
            print(f"  {m}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Map SFXL + prior Gold Sheet (PGS) into Gold Sheet 2.0 PDF fields."
    )
    p.add_argument("--sfxl", type=Path, help="Builder Accounts Salesforce Report (.xlsx)")
    p.add_argument("--pgs", type=Path, help="Prior gold sheet (.xlsx), e.g. GS - Shea Homes 2026")
    p.add_argument("--pdf", type=Path, help="Gold Sheet 2.0 Internal Version fillable PDF")
    p.add_argument(
        "--out",
        type=Path,
        help="Output filled PDF path (default: same as PGS name with .pdf, e.g. GS - Shea Homes 2026.pdf)",
    )
    p.add_argument("--sfxl-sheet", default=None, help="Optional SFXL sheet name")
    p.add_argument("--pgs-sheet", default=None, help="Optional PGS sheet name")
    p.add_argument(
        "--account",
        default=None,
        help="Account/builder name to select in SFXL (defaults to PGS Account Name)",
    )
    p.add_argument(
        "--all-accounts",
        action="store_true",
        help="Do not filter SFXL to Corporate Account Names (default: Corporate only)",
    )
    p.add_argument(
        "--mapping-json",
        type=Path,
        default=None,
        help="Also write resolved logical->value mapping JSON",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Resolve mapping and print it; do not write PDF",
    )
    p.add_argument("--inspect-pdf", type=Path, help="List AcroForm field names and exit")
    p.add_argument("--inspect-pgs", type=Path, help="Dump PGS labels/strategies/divisions and exit")
    p.add_argument("--inspect-sfxl", type=Path, help="Dump SFXL headers/sample and exit")
    return p.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    corporate_only = not args.all_accounts

    if args.inspect_pdf:
        inspect_pdf(args.inspect_pdf)
        return 0
    if args.inspect_pgs:
        inspect_pgs(args.inspect_pgs, args.pgs_sheet)
        return 0
    if args.inspect_sfxl:
        inspect_sfxl(args.inspect_sfxl, args.sfxl_sheet, corporate_only=corporate_only)
        return 0

    if not args.sfxl or not args.pgs or not args.pdf:
        print(
            "ERROR: --sfxl, --pgs, and --pdf are required unless using an --inspect-* flag.",
            file=sys.stderr,
        )
        return 2

    for path, label in (
        (args.sfxl, "SFXL"),
        (args.pgs, "PGS"),
        (args.pdf, "PDF"),
    ):
        if not path.exists():
            print(f"ERROR: {label} not found: {path}", file=sys.stderr)
            return 2

    pgs = load_pgs(args.pgs, args.pgs_sheet)
    # Full report for Divisions_*; Corporate subset for HQ / Website / Type match.
    headers, sfxl_all_rows, colmap = load_sfxl(
        args.sfxl, args.sfxl_sheet, corporate_only=False
    )
    corp_rows = [
        r
        for r in sfxl_all_rows
        if is_corporate_account_name(sfxl_get(r, "Account Name", colmap, headers))
    ]
    print(
        f"SFXL: {len(sfxl_all_rows)} total account row(s); "
        f"{len(corp_rows)} Corp account(s) available for HQ match."
    )

    account = args.account or pgs_lookup(pgs, "Account Name")
    if args.all_accounts:
        sfxl_row = pick_sfxl_row(sfxl_all_rows, headers, colmap, account)
    else:
        sfxl_row = pick_sfxl_row(
            corp_rows,
            headers,
            colmap,
            account,
            fallback_rows=sfxl_all_rows,
        )
    if sfxl_row is None:
        return 3

    values = build_field_values(
        pgs, sfxl_row, headers, colmap, sfxl_all_rows=sfxl_all_rows
    )

    print("\n=== Resolved field mapping ===")
    for k in PDF_FIELD_CANDIDATES:
        v = values.get(k, "")
        preview = v if len(v) <= 160 else v[:157] + "..."
        print(f"{k}: {preview}")
    print(f"\nDate source: today's date (run date) = {values.get('Date', '')}")

    if args.mapping_json:
        args.mapping_json.parent.mkdir(parents=True, exist_ok=True)
        args.mapping_json.write_text(json.dumps(values, indent=2), encoding="utf-8")
        print(f"\nWrote mapping JSON: {args.mapping_json}")

    if args.dry_run:
        print("\nDry run only — PDF not written.")
        return 0

    out = resolve_output_path(args.pgs, args.out)
    print(f"Output PDF: {out}")
    written = fill_pdf(args.pdf, values, out)
    print(f"\nWrote PDF: {out}")
    print(f"PDF fields updated: {len(written)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
