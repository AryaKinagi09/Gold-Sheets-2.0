"""Single point of contact for builder data.

Routes and fill logic call these methods only.
MongoDataAccess is the current implementation.
A SalesforceDataAccess class will use the same methods later.
Switch with DATA_SOURCE=mongo or DATA_SOURCE=salesforce.
"""

from __future__ import annotations

import os
from typing import Any, Protocol

from pymongo.database import Database


class DataAccess(Protocol):
    def list_builders(self) -> list[dict[str, Any]]: ...

    def list_regions(self, builder_key: str) -> list[dict[str, Any]]: ...

    def get_prior_gold_sheet(
        self, builder_key: str, year: int = 2026
    ) -> dict[str, Any] | None: ...

    def get_sfxl_accounts(
        self,
        builder_key: str,
        region_labels: list[str] | None = None,
    ) -> list[dict[str, Any]]: ...


class MongoDataAccess:
    def __init__(self, db: Database):
        self.db = db

    def list_builders(self) -> list[dict[str, Any]]:
        rows = list(
            self.db.sfxl_accounts.aggregate(
                [
                    {
                        "$group": {
                            "_id": "$builder_key",
                            "corporate_names": {
                                "$push": {
                                    "$cond": [
                                        {"$eq": ["$is_corporate", True]},
                                        "$account_name",
                                        "$$REMOVE",
                                    ]
                                }
                            },
                            "region_count": {
                                "$sum": {
                                    "$cond": [
                                        {
                                            "$and": [
                                                {"$eq": ["$is_corporate", False]},
                                                {
                                                    "$gt": [
                                                        {
                                                            "$strLenCP": {
                                                                "$ifNull": [
                                                                    "$region_label",
                                                                    "",
                                                                ]
                                                            }
                                                        },
                                                        0,
                                                    ]
                                                },
                                            ]
                                        },
                                        1,
                                        0,
                                    ]
                                }
                            },
                        }
                    },
                    {"$sort": {"_id": 1}},
                ]
            )
        )
        priors = {
            doc["builder_key"]: doc.get("builder_name")
            for doc in self.db.prior_gold_sheets.find(
                {}, {"_id": 0, "builder_key": 1, "builder_name": 1}
            )
        }

        builders = []
        for row in rows:
            key = row["_id"]
            if not key:
                continue
            corporate_names = row.get("corporate_names") or []
            builders.append(
                {
                    "builder_key": key,
                    "builder_name": priors.get(key)
                    or (corporate_names[0] if corporate_names else None)
                    or key.replace("_", " ").title(),
                    "region_count": row.get("region_count", 0),
                    "has_prior_gold_sheet": key in priors,
                }
            )
        return builders

    def list_regions(self, builder_key: str) -> list[dict[str, Any]]:
        return list(
            self.db.sfxl_accounts.find(
                {
                    "builder_key": builder_key,
                    "is_corporate": False,
                    "region_label": {"$nin": [None, ""]},
                },
                {"_id": 0, "region_label": 1, "account_name": 1, "physical_state": 1},
            ).sort("region_label", 1)
        )

    def get_prior_gold_sheet(
        self, builder_key: str, year: int = 2026
    ) -> dict[str, Any] | None:
        return self.db.prior_gold_sheets.find_one(
            {"builder_key": builder_key, "year": year},
            {"_id": 0},
        )

    def get_sfxl_accounts(
        self,
        builder_key: str,
        region_labels: list[str] | None = None,
    ) -> list[dict[str, Any]]:
        docs = list(
            self.db.sfxl_accounts.find({"builder_key": builder_key}, {"_id": 0})
        )
        if region_labels is None:
            return docs
        wanted = set(region_labels)
        return [
            doc
            for doc in docs
            if doc.get("is_corporate") or doc.get("region_label") in wanted
        ]


def get_data_access(db: Database) -> DataAccess:
    source = os.environ.get("DATA_SOURCE", "mongo")
    if source == "mongo":
        return MongoDataAccess(db)
    raise RuntimeError(f"Unsupported DATA_SOURCE: {source}")