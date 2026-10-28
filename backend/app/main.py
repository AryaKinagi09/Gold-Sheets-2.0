import os
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from pymongo import MongoClient

from app.data_access import get_data_access
from app.generate import (
    build_field_values_from_mongo,
    generate_division_bundle,
    generate_pdf,
    safe_filename,
)

app = FastAPI(title="Gold Sheets API")


def get_db():
    client = MongoClient(
        os.environ["MONGO_URI"],
        serverSelectionTimeoutMS=3000,
    )
    return client[os.environ.get("MONGO_DB", "gold_sheets")]


class GenerateRequest(BaseModel):
    builder_key: str
    sheet_type: Literal["corporate", "division"] = "corporate"
    regions: list[str] | None = Field(
        default=None,
        description=(
            "Division labels. Corporate ignores this and always uses all "
            "divisions. Division mode requires one or more labels."
        ),
    )
    year: int = 2026
    dry_run: bool = False


@app.get("/health")
def health():
    db = get_db()
    db.client.admin.command("ping")
    return {
        "status": "ok",
        "sfxl_accounts": db.sfxl_accounts.count_documents({}),
        "prior_gold_sheets": db.prior_gold_sheets.count_documents({}),
    }


@app.get("/builders")
def builders():
    return get_data_access(get_db()).list_builders()


@app.get("/divisions")
def divisions():
    """Operating divisions only (no corporate accounts)."""
    return get_data_access(get_db()).list_divisions()


@app.get("/builders/{builder_key}/regions")
def regions(builder_key: str):
    data = get_data_access(get_db())
    known = {item["builder_key"] for item in data.list_builders()}
    if builder_key not in known:
        raise HTTPException(status_code=404, detail="Builder not found")
    return data.list_regions(builder_key)


@app.post("/generate")
def generate(body: GenerateRequest):
    data = get_data_access(get_db())
    prior = data.get_prior_gold_sheet(body.builder_key, body.year)
    if not prior:
        raise HTTPException(
            status_code=404,
            detail=f"No prior Gold Sheet for builder_key={body.builder_key}",
        )

    # Corporate always loads every account row so Divisions_* lists all units.
    # Division mode still loads all rows for corporate header fields; the
    # selected labels only control which division PDF(s) are produced.
    sfxl_docs = data.get_sfxl_accounts(body.builder_key, None)
    if not sfxl_docs:
        raise HTTPException(
            status_code=404,
            detail=f"No SFXL accounts for builder_key={body.builder_key}",
        )

    if body.sheet_type == "division":
        if not body.regions:
            raise HTTPException(
                status_code=400,
                detail="Select at least one division for a division sheet.",
            )

    if body.dry_run:
        if body.sheet_type == "corporate":
            field_values = build_field_values_from_mongo(
                prior, sfxl_docs, sheet_type="corporate"
            )
            return {
                "builder_key": body.builder_key,
                "sheet_type": body.sheet_type,
                "regions": None,
                "pages": [1, 2],
                "field_values": field_values,
            }
        return {
            "builder_key": body.builder_key,
            "sheet_type": body.sheet_type,
            "regions": body.regions,
            "pages": [3, 4],
            "field_values": [
                build_field_values_from_mongo(
                    prior,
                    sfxl_docs,
                    sheet_type="division",
                    division_label=label,
                )
                for label in (body.regions or [])
            ],
        }

    template = Path(os.environ.get("PDF_TEMPLATE", "/data/template.pdf"))
    if not template.exists():
        raise HTTPException(
            status_code=500,
            detail=f"PDF template not found at {template}",
        )

    out_dir = Path(os.environ.get("OUTPUT_DIR", "/data/outputs"))
    out_dir.mkdir(parents=True, exist_ok=True)
    builder_name = prior.get("builder_name") or body.builder_key

    if body.sheet_type == "corporate":
        out_name = f"GS_{safe_filename(builder_name)}_{body.year}_corporate.pdf"
        out_path = out_dir / out_name
        generate_pdf(
            prior=prior,
            sfxl_docs=sfxl_docs,
            template_path=template,
            output_path=out_path,
            sheet_type="corporate",
        )
        return FileResponse(
            path=out_path,
            media_type="application/pdf",
            filename=out_name,
        )

    out_path = generate_division_bundle(
        prior=prior,
        sfxl_docs=sfxl_docs,
        division_labels=body.regions or [],
        template_path=template,
        output_dir=out_dir,
        year=body.year,
    )
    media = (
        "application/zip"
        if out_path.suffix.lower() == ".zip"
        else "application/pdf"
    )
    return FileResponse(
        path=out_path,
        media_type=media,
        filename=out_path.name,
    )


frontend_dir = Path(os.environ.get("FRONTEND_DIR", "/app/frontend"))
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
