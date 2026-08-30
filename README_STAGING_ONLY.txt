VAT Report STAGING Update
=========================

Scope
-----
Modify only accounting_report.html on a new branch based on staging.
Do NOT touch main or Production.

Recommended branch:
feat/staging-vat-report-layout-20260830

Expected VAT report:
1. Separate "مخرجات ضريبة القيمة المضافة" table.
2. Separate "مدخلات ضريبة القيمة المضافة" table.
3. Columns:
   party name, VAT number, invoice number, date, amount before VAT, VAT amount.
4. Totals for each section.
5. VAT summary:
   output VAT, input VAT, net VAT = output - input.
6. Existing filters, print/PDF and Excel link remain available.
7. Non-VAT accounting reports remain unchanged.

Acceptance test on STAGING
--------------------------
- Open report 15 "ضريبة القيمة المضافة".
- Existing posted sales invoice INV-2026-000002 should appear under outputs.
- Expected base amount: 180.00 SAR
- Expected VAT: 27.00 SAR
- It must NOT appear under inputs.
- Default status remains "مرحّل".
- Date filters must work.
- Print/PDF must show both sections and the net VAT summary.
- Excel export must remain available and export VAT data.
