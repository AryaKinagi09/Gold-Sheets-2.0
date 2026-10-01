"""Build Gold Sheet field values from Mongo (via DataAccess) and fill the PDF."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

from fill_gold_sheet import (
    SfxlRow,
    _map_columns,
    build_divisions_from_sfxl,
    fill_pdf,
    format_currency_usd,
    format_integer_commas,
    format_spec_percent,
    is_corporate_account_name,
    normalize_display_casing,
    sfxl_get,
)


def _mongo_docs_to_sfxl(
    docs: list[dict[str, Any]],
) -> tuple[list[str], list[SfxlRow], dict[str, int]]:
    rows: list[SfxlRow] = []
    headers: list[str] = []
    for doc in docs:
        raw = doc.get("raw") or {}
        if not raw:
            continue
        if not headers:
            headers = list(raw.keys())
        rows.append(SfxlRow(raw=raw))
    colmap = _map_columns(headers) if headers else {}
    return headers, rows, colmap


def _pick_corporate_row(
    rows: list[SfxlRow], headers: list[str], colmap: dict[str, int]
) -> SfxlRow | None:
    for row in rows:
        name = sfxl_get(row, "Account Name", colmap, headers)
        if is_corporate_account_name(name):
            return row
    return rows[0] if rows else None


def build_field_values_from_mongo(
    prior: dict[str, Any],
    sfxl_docs: list[dict[str, Any]],
) -> dict[str, str]:
    fields = prior.get("fields") or {}
    builder_name = (
        fields.get("Account Name")
        or prior.get("builder_name")
        or ""
    )

    headers, rows, colmap = _mongo_docs_to_sfxl(sfxl_docs)
    corp = _pick_corporate_row(rows, headers, colmap)

    def sf(logical: str) -> str:
        if not corp:
            return ""
        return sfxl_get(corp, logical, colmap, headers)

    strategies = fields.get("Builder Strategy") or []
    if isinstance(strategies, str):
        strategies = [strategies]
    strategies = [normalize_display_casing(s) for s in strategies[:4]]
    while len(strategies) < 4:
        strategies.append("")

    div_buckets, total_d, _ = build_divisions_from_sfxl(
        rows, headers, colmap, builder_name, max_buckets=4
    )

    return {
        "Builder_Name": builder_name,
        "Headquarters": sf("Physical State/Province"),
        "Website_URL": sf("Website"),
        "Market_Facing_Position": sf("Type"),
        "Date": date.today().strftime("%m/%d/%Y"),
        "NAM_Name": fields.get("National Account Manager") or "",
        "NAM_Email": fields.get("NAM Email") or "",
        "Builder_Rank": fields.get("Builder Rank") or "",
        "Annual_Closings": format_integer_commas(
            fields.get("Annual Closings") or ""
        ),
        "National_ASP": format_currency_usd(
            fields.get("National Average Selling Price") or ""
        ),
        "Spec_Percent": format_spec_percent(fields.get("Spec Home %") or ""),
        "Rebate_Structure": normalize_display_casing(
            fields.get("Rebate Structure") or ""
        ),
        "Model_Home_Design_Center_Program": normalize_display_casing(
            fields.get("Design Center Program") or ""
        ),
        "Pricing_Actions": normalize_display_casing(
            fields.get("2026 Pricing Action Status") or ""
        ),
        "National_Pricing": normalize_display_casing(
            fields.get("National Pricing") or ""
        ),
        "Special_Terms": normalize_display_casing(
            fields.get("Special Terms") or ""
        ),
        "GrowthStrat1_Text": strategies[0],
        "GrowthStrat2_Text": strategies[1],
        "GrowthStrat3_Text": strategies[2],
        "GrowthStrat4_Text": strategies[3],
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


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^\w\-]+", "_", name.strip())
    return cleaned.strip("_") or "gold_sheet"


def generate_pdf(
    *,
    prior: dict[str, Any],
    sfxl_docs: list[dict[str, Any]],
    template_path: Path,
    output_path: Path,
) -> dict[str, str]:
    values = build_field_values_from_mongo(prior, sfxl_docs)
    fill_pdf(template_path, values, output_path)
    return values