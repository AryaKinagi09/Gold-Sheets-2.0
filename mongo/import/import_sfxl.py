#!/usr/bin/env python3
"""Import Builder Accounts Salesforce Report (.xlsx) into MongoDB."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pymongo import MongoClient, UpdateOne

# Allow importing fill_gold_sheet from project root
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fill_gold_sheet import (  # noqa: E402
    builder_core_name,
    is_corporate_account_name,
    load_sfxl,
    sfxl_get,
    strip_builder_from_account_name,
)


def parent_builder_name(account_name: str) -> str:
    """
    Parent builder identity for a Corporate or region row.
    'Shea Homes - CA Central Coast' -> 'shea homes'
    'Shea Homes Corporate' -> 'shea homes'
    """
    if " - " in account_name:
        left = account_name.split(" - ", 1)[0].strip()
        return builder_core_name(left) or left
    return builder_core_name(account_name) or account_name


def builder_key_from_name(account_name: str) -> str:
    core = parent_builder_name(account_name)
    return core.lower().replace(" ", "_")


def main() -> int:
    p = argparse.ArgumentParser(description="Import SFXL workbook into MongoDB")
    p.add_argument(
        "--sfxl",
        type=Path,
        default=ROOT / "Builder Accounts Salesforce Report.xlsx",
        help="Path to Salesforce report .xlsx",
    )
    p.add_argument(
        "--mongo-uri",
        default="mongodb://goldsheets:goldsheets_dev@localhost:27017/?authSource=admin",
    )
    p.add_argument("--db", default="gold_sheets")
    p.add_argument("--collection", default="sfxl_accounts")
    args = p.parse_args()

    if not args.sfxl.exists():
        print(f"ERROR: file not found: {args.sfxl}")
        return 1

    # ALL rows (Corporate + divisions/regions)
    headers, rows, colmap = load_sfxl(args.sfxl, corporate_only=False)
    print(f"Loaded {len(rows)} SFXL row(s) from {args.sfxl.name}")

    client = MongoClient(args.mongo_uri)
    col = client[args.db][args.collection]

    ops = []
    for row in rows:
        account_name = sfxl_get(row, "Account Name", colmap, headers)
        if not account_name.strip():
            continue

        bkey = builder_key_from_name(account_name)
        corporate = is_corporate_account_name(account_name)
        region_label = ""
        if not corporate:
            parent = parent_builder_name(account_name)
            region_label = strip_builder_from_account_name(account_name, parent)
            if not region_label and " - " in account_name:
                region_label = account_name.split(" - ", 1)[1].strip()

        doc = {
            "builder_key": bkey,
            "account_name": account_name,
            "is_corporate": corporate,
            "region_label": region_label or None,
            "website": sfxl_get(row, "Website", colmap, headers),
            "type": sfxl_get(row, "Type", colmap, headers),
            "physical_state": sfxl_get(row, "Physical State/Province", colmap, headers),
            "last_modified_date": sfxl_get(row, "Last Modified Date", colmap, headers),
            # Keep full row so fill logic can use any column later
            "raw": row.raw,
            "source": "sfxl_xlsx",
        }

        ops.append(
            UpdateOne(
                {"account_name": account_name},
                {"$set": doc},
                upsert=True,
            )
        )

    if not ops:
        print("No rows to import.")
        return 1

    result = col.bulk_write(ops)
    col.create_index("builder_key")
    col.create_index("account_name", unique=True)
    col.create_index([("builder_key", 1), ("is_corporate", 1)])

    print(
        f"Done. upserted={result.upserted_count}, "
        f"modified={result.modified_count}, matched={result.matched_count}"
    )
    print(f"Total docs in {args.db}.{args.collection}: {col.count_documents({})}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())