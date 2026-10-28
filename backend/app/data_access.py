"""Single point of contact for builder data.

Routes and fill logic call these methods only.
MongoDataAccess is the current implementation.
A SalesforceDataAccess class will use the same methods later.
Switch with DATA_SOURCE=mongo or DATA_SOURCE=salesforce.
"""

from __future__ import annotations

import os
import re
from typing import Any, Protocol

from pymongo.database import Database

from fill_gold_sheet import builder_core_name, is_corporate_account_name


_HQ_LABELS = frozenset({"hq", "corporate", "corp", "headquarters"})
_CORPORATE_IN_NAME = re.compile(
    r"\b(?:corporate|corp|headquarters|hq)\b",
    re.IGNORECASE,
)


def _display_builder_name(raw: str, fallback_key: str = "") -> str:
    """Human builder name without Corporate/HQ suffixes."""
    core = builder_core_name(raw or "") or builder_core_name(
        (fallback_key or "").replace("_", " ")
    )
    if not core:
        return (raw or fallback_key or "").strip()
    return " ".join(part.capitalize() for part in core.split())


def _is_hq_or_corporate_label(label: str) -> bool:
    text = (label or "").strip().lower()
    if not text:
        return True
    if text in _HQ_LABELS:
        return True
    return bool(_CORPORATE_IN_NAME.search(text))


class DataAccess(Protocol):
    def list_builders(self) -> list[dict[str, Any]]: ...

    def list_divisions(self) -> list[dict[str, Any]]: ...

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
        """One entry per builder_key. Kept simple so the API never returns empty on agg errors."""
        priors = {
            doc["builder_key"]: doc.get("builder_name")
            for doc in self.db.prior_gold_sheets.find(
                {}, {"_id": 0, "builder_key": 1, "builder_name": 1}
            )
        }

        by_key: dict[str, dict[str, Any]] = {}
        cursor = self.db.sfxl_accounts.find(
            {},
            {
                "_id": 0,
                "builder_key": 1,
                "account_name": 1,
                "is_corporate": 1,
                "region_label": 1,
            },
        )
        for doc in cursor:
            key = doc.get("builder_key")
            if not key:
                continue
            entry = by_key.setdefault(
                key,
                {
                    "corporate_names": [],
                    "region_count": 0,
                },
            )
            account_name = str(doc.get("account_name") or "").strip()
            label = str(doc.get("region_label") or "").strip()
            is_corp = bool(doc.get("is_corporate")) or (
                is_corporate_account_name(account_name) and " - " not in account_name
            )
            if is_corp and account_name:
                entry["corporate_names"].append(account_name)
            elif label and not _is_hq_or_corporate_label(label):
                entry["region_count"] += 1

        builders = []
        for key in sorted(by_key.keys()):
            entry = by_key[key]
            corporate_names = entry["corporate_names"]
            raw_name = (
                priors.get(key)
                or (corporate_names[0] if corporate_names else None)
                or key.replace("_", " ")
            )
            try:
                builder_name = _display_builder_name(str(raw_name), key)
            except Exception:
                builder_name = str(raw_name).strip() or key.replace("_", " ").title()
            builders.append(
                {
                    "builder_key": key,
                    "builder_name": builder_name,
                    "region_count": entry["region_count"],
                    "has_corporate": len(corporate_names) > 0,
                    "has_prior_gold_sheet": key in priors,
                }
            )
        return builders

    def list_divisions(self) -> list[dict[str, Any]]:
        """All operating divisions across builders (never corporate / HQ rows)."""
        priors = {
            doc["builder_key"]: {
                "builder_name": doc.get("builder_name"),
                "has_prior": True,
            }
            for doc in self.db.prior_gold_sheets.find(
                {}, {"_id": 0, "builder_key": 1, "builder_name": 1}
            )
        }
        builders = {b["builder_key"]: b for b in self.list_builders()}
        rows = list(
            self.db.sfxl_accounts.find(
                {
                    "region_label": {"$nin": [None, ""]},
                },
                {
                    "_id": 0,
                    "builder_key": 1,
                    "region_label": 1,
                    "account_name": 1,
                    "physical_state": 1,
                    "is_corporate": 1,
                },
            ).sort([("builder_key", 1), ("region_label", 1)])
        )
        divisions = []
        for row in rows:
            account_name = str(row.get("account_name") or "").strip()
            label = str(row.get("region_label") or "").strip()
            # Keep market divisions even when SF names them
            # "Builder Corporate - Arizona". Drop HQ-only / corporate-only rows.
            if not label or _is_hq_or_corporate_label(label):
                continue
            if is_corporate_account_name(account_name) and " - " not in account_name:
                continue
            key = row.get("builder_key") or ""
            builder = builders.get(key) or {}
            prior = priors.get(key) or {}
            builder_name = _display_builder_name(
                str(
                    prior.get("builder_name")
                    or builder.get("builder_name")
                    or account_name
                    or key.replace("_", " ")
                ),
                key,
            )
            divisions.append(
                {
                    "builder_key": key,
                    "builder_name": builder_name,
                    "region_label": label,
                    "account_name": account_name,
                    "physical_state": row.get("physical_state") or "",
                    "display_name": f"{builder_name} - {label}",
                    "has_prior_gold_sheet": bool(prior.get("has_prior")),
                }
            )
        return divisions

    def list_regions(self, builder_key: str) -> list[dict[str, Any]]:
        # Operating divisions only — never corporate / HQ rows.
        rows = list(
            self.db.sfxl_accounts.find(
                {
                    "builder_key": builder_key,
                    "region_label": {"$nin": [None, ""]},
                },
                {
                    "_id": 0,
                    "region_label": 1,
                    "account_name": 1,
                    "physical_state": 1,
                    "is_corporate": 1,
                },
            ).sort("region_label", 1)
        )
        cleaned = []
        for row in rows:
            account_name = str(row.get("account_name") or "").strip()
            label = str(row.get("region_label") or "").strip()
            if not label or _is_hq_or_corporate_label(label):
                continue
            if is_corporate_account_name(account_name) and " - " not in account_name:
                continue
            cleaned.append(
                {
                    "region_label": label,
                    "account_name": account_name,
                    "physical_state": row.get("physical_state") or "",
                }
            )
        return cleaned

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