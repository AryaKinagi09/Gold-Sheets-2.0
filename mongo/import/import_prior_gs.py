#!/usr/bin/env python3
"""Import Prior Gold Sheet workbooks into MongoDB as structured field docs."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pymongo import MongoClient, UpdateOne

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from fill_gold_sheet import (  # noqa: E402
    builder_core_name,
    extract_builder_strategies,
    extract_label_block_joined,
    extract_nam_from_pgs,
    load_pgs,
    pgs_lookup,
)


def builder_key_from_name(name: str) -> str:
    core = builder_core_name(name) or name
    return core.lower().replace(" ", "_")


def extract_prior_fields(pgs) -> dict:
    strategies = extract_builder_strategies(pgs, max_items=4)
    nam_name, nam_email = extract_nam_from_pgs(pgs)
    return {
        "Account Name": pgs_lookup(pgs, "Account Name"),
        "Builder Rank": pgs_lookup(pgs, "Builder Rank"),
        "Annual Closings": pgs_lookup(pgs, "Annual Closings"),
        "National Average Selling Price": pgs_lookup(
            pgs, "National Average Selling Price"
        ),
        "Spec Home %": pgs_lookup(pgs, "Spec Home %"),
        "Rebate Structure": extract_label_block_joined(pgs, "Rebate Structure")
        or pgs_lookup(pgs, "Rebate Structure"),
        "Design Center Program": extract_label_block_joined(pgs, "Design Center Program")
        or pgs_lookup(pgs, "Design Center Program"),
        "2026 Pricing Action Status": extract_label_block_joined(
            pgs, "2026 Pricing Action Status"
        )
        or pgs_lookup(pgs, "2026 Pricing Action Status"),
        "National Pricing": pgs_lookup(pgs, "National Pricing"),
        "Special Terms": pgs_lookup(pgs, "Special Terms"),
        "Builder Strategy": strategies,
        "Operating Divisions": pgs_lookup(pgs, "Operating Divisions"),
        "National Account Manager": nam_name,
        "NAM Email": nam_email,
    }


def main() -> int:
    p = argparse.ArgumentParser(description="Import Prior GS workbooks into MongoDB")
    p.add_argument(
        "--files",
        nargs="*",
        type=Path,
        help="Specific .xlsx/.xlsm files. If omitted, imports all GS - *.xlsx/xlsm in project root.",
    )
    p.add_argument(
        "--mongo-uri",
        default="mongodb://goldsheets:goldsheets_dev@localhost:27018/?authSource=admin",
    )
    p.add_argument("--db", default="gold_sheets")
    p.add_argument("--collection", default="prior_gold_sheets")
    p.add_argument("--year", type=int, default=2026)
    args = p.parse_args()

    if args.files:
        files = args.files
    else:
        files = sorted(ROOT.glob("GS - *.xlsx")) + sorted(ROOT.glob("GS - *.xlsm"))

    if not files:
        print("ERROR: no Prior GS files found.")
        return 1

    client = MongoClient(args.mongo_uri)
    col = client[args.db][args.collection]
    ops = []

    for path in files:
        if not path.exists():
            print(f"SKIP missing: {path}")
            continue
        print(f"Reading {path.name} ...")
        pgs = load_pgs(path)
        fields = extract_prior_fields(pgs)
        builder_name = fields.get("Account Name") or path.stem
        bkey = builder_key_from_name(builder_name)

        doc = {
            "builder_key": bkey,
            "builder_name": builder_name,
            "year": args.year,
            "source_file": path.name,
            "fields": fields,
            "source": "prior_gs_xlsx",
        }
        ops.append(
            UpdateOne(
                {"builder_key": bkey, "year": args.year},
                {"$set": doc},
                upsert=True,
            )
        )
        print(f"  -> builder_key={bkey}, name={builder_name}")

    if not ops:
        print("ERROR: nothing imported.")
        return 1

    result = col.bulk_write(ops)
    col.create_index([("builder_key", 1), ("year", 1)], unique=True)

    print(
        f"Done. upserted={result.upserted_count}, "
        f"modified={result.modified_count}, matched={result.matched_count}"
    )
    print(f"Total docs in {args.db}.{args.collection}: {col.count_documents({})}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())