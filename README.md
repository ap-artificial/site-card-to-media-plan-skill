# Site-Card Deck → Media Plan Excel Skill

Final step of enterprise OOH workflow. Converts site-card decks into client-ready Media Plan Excel.

Why deterministic
Input is fixed-layout deck from agency template or clean sites.json, so template-filling is safe.

Workflow
1. Get per-site data: sites.json > parse deck via parse_deck.py
2. Get campaign fields: Client, campaign, dates, currency (ask user!)
3. Build Excel: build_xlsx.py sites.json campaign.json template.xlsx output.xlsx
4. QA: convert to pdf -> check VAT, TOTAL COST, currency

Key Details
- Template: Use client Media Plan template (no proprietary file committed)
- Starts row 21, auto-inserts rows
- Formulas: VAT = Media Cost*11%, TOTAL = Media + VAT + Production
- QA critical: Price format must match locale

Pipeline
vendor-proposal -> site-card deck -> Media Plan Excel (This Skill)

Stack: python-pptx, openpyxl, soffice, pdftotext
No client data included.
