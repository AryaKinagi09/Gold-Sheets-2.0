# Gold Sheets 2.0

Internal tool to generate **Gold Sheet 2.0 (Internal)** PDFs from builder account data and prior Gold Sheet content.

## Goal

Reduce manual assembly of Gold Sheets by automating field fill from known sources, with a clear path from a local script → containerized app → Salesforce-backed data → AI-assisted content.

## Phases

### P1 — Script (current foundation)
- Local Python script (`fill_gold_sheet.py`)
- Inputs:
  - **SFXL**: Builder Accounts Salesforce Report (`.xlsx`)
  - **PGS**: Prior Gold Sheet workbook (`.xlsx` / `.xlsm`)
  - **Template**: Gold Sheet 2.0 Internal fillable PDF
- Output: filled PDF (+ optional JSON preview)
- Status: working for multiple builders; QA / edge cases still manual

### P2 — Containerized application (in progress)
- Dockerized stack: API + simple frontend + database
- **MongoDB stand-ins** for the two generation sources until Salesforce access is finalized:
  1. Salesforce-like account / division rows (SFXL)
  2. Prior Gold Sheet–like structured fields (PGS)
- UI: select a **builder**, then one / many / **all regions** under that builder; generate Gold Sheet(s)
- Reuses existing fill logic from P1 (adapted to read Mongo instead of Excel where needed)
- Later in P2: replace Mongo stand-ins with real Salesforce pull (query or report + API), per IT guidance

### P3 — AI integration (planned)
- Assist with collection, summarization, and inference for fields that are not direct lookups
- Human review / QA remains part of the production process

## Architecture (P2 target)

| Service    | Role |
|-----------|------|
| Frontend  | Builder + region selection; trigger generate; download results |
| API       | Orchestrates generation using P1 logic |
| MongoDB   | Temporary stand-in for SFXL + Prior GS sources |
| PDF template | Packaged with the app or mounted as a volume |

## Local development (P2)

*(Commands will be filled in as Docker Compose is added.)*

```bash
docker compose up --build