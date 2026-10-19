# Billups Site-Card Deck → Media Plan Excel Skill

Final step of the billups OOH workflow. Converts site-card decks into client-ready Media Plan Excel.

### Why this is deterministic
Unlike vendor-proposal step (messy input), here input is fixed-layout deck from `billups site template.pptx` or clean `sites.json`. So we can use template-filling script safely.

### Workflow
1.  **Get per-site data:** Prefer `sites.json` (clean numeric fields) > parse deck via `python3 scripts/parse_deck.py <deck.pptx> sites.json`
2.  **Get campaign fields:** Client name, campaign name, dates, currency (ask user, don't guess!)
3.  **Build Excel:** `python3 scripts/build_xlsx.py sites.json campaign.json Media_Plan_Template_billups_v3.xlsx output.xlsx`
4.  **QA:** `soffice --headless --convert-to pdf` → `pdftotext -layout` → check formulas, VAT, TOTAL COST, currency symbol

### Key Details
- Template: `Media_Plan_Template_billups_v3.xlsx` (client-facing, NOT internal app export)
- Starts row 21, 8 pre-formatted rows, auto-inserts more if needed
- Formulas: VAT = Media Cost*11%, TOTAL COST = Media Cost + VAT + Production
- Columns: W=Direct, AE=1 month, Z=SOT 100% Static / 1/N Digital
- QA critical: Price format "360.000.000,-" vs "360,000,000" can cause 1000x error!

### Pipeline
`vendor-proposal → site-card deck (Skill 2) → Media Plan Excel (This Skill)`

### Stack
`python-pptx, openpyxl, soffice, pdftotext`
