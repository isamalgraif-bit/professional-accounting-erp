#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from pathlib import Path
import re
import subprocess
import sys

EXPECTED_BRANCH = "1-staging-2140---fix-critical-defects-and-retest-erp"
APP = Path("app.py")

def run(*args, check=True):
    p = subprocess.run(args, text=True, capture_output=True)
    if check and p.returncode != 0:
        print(p.stdout)
        print(p.stderr, file=sys.stderr)
        raise SystemExit(p.returncode)
    return p

def replace_once(text, old, new, label):
    count = text.count(old)
    if count != 1:
        raise RuntimeError(f"{label}: expected exactly 1 match, found {count}")
    return text.replace(old, new, 1)

branch = run("git", "branch", "--show-current").stdout.strip()
if branch != EXPECTED_BRANCH:
    raise SystemExit(
        f"SAFETY STOP: current branch is {branch!r}, expected {EXPECTED_BRANCH!r}. No files were changed."
    )

if not APP.exists():
    raise SystemExit("app.py not found. Run this script from the repository root.")

source = APP.read_text(encoding="utf-8")
original = source

source = replace_once(
    source,
    'expenses_total = row("SELECT COALESCE(SUM(total),0) s FROM expenses")["s"]',
    'expenses_total = row("SELECT COALESCE(SUM(amount),0) s FROM expenses")["s"]',
    "dashboard expenses total",
)

source = replace_once(
    source,
    '            value = request.form.get(field, "")',
    '            value = request.form.get(field, record.get(field) if field in ("phone", "email") else "")',
    "generic customer contact preservation",
)

party_start = source.find("def party_edit(")
if party_start != -1:
    party_end = source.find("\n@app.route", party_start + 1)
    if party_end != -1:
        party = source[party_start:party_end]
        party = re.sub(
            r'(\bphone\s*=\s*)\(request\.form\.get\("phone"\)\s*or\s*""\)\.strip\(\)',
            r'\1(request.form.get("phone", party.get("phone") or "") or "").strip()',
            party,
            count=1,
        )
        party = re.sub(
            r'(\bemail\s*=\s*)\(request\.form\.get\("email"\)\s*or\s*""\)\.strip\(\)',
            r'\1(request.form.get("email", party.get("email") or "") or "").strip()',
            party,
            count=1,
        )
        source = source[:party_start] + party + source[party_end:]

item_old = """    if entity=="item":
        code=next_inventory_sku()
        existing=row("SELECT id FROM inventory WHERE code=:code OR name=:name",
                     {"code":code,"name":name})
        if existing:
            raise ValueError("المادة موجودة مسبقًا.")
        execute("""INSERT INTO inventory(code,sku,name,description,unit,quantity,unit_cost,cost,
          reorder_level,active)
          VALUES(:code,:code,:name,:description,:unit,:quantity,:cost,:cost,:reorder,1)""",
          {"code":code,"name":name,"description":payload.get("description",""),
           "unit":normalize_item_unit(payload.get("unit")),
           "quantity":float(payload.get("quantity") or 0),
           "cost":float(payload.get("unit_cost") or 0),
           "reorder":float(payload.get("reorder_level") or 0)})
        new_id=row("SELECT id FROM inventory WHERE code=:code",{"code":code})["id"]
        return "items",smart_entity_row("items",new_id)"""

item_new = """    if entity=="item":
        code=next_inventory_sku()
        existing=row("SELECT id FROM inventory WHERE code=:code OR name=:name",
                     {"code":code,"name":name})
        if existing:
            raise ValueError("المادة موجودة مسبقًا.")
        opening_quantity=round(float(payload.get("quantity") or 0),3)
        cost=round(float(payload.get("unit_cost") or 0),4)
        if opening_quantity < 0:
            raise ValueError("الرصيد الافتتاحي لا يمكن أن يكون سالبًا.")
        warehouse=None
        if opening_quantity > 0:
            requested_warehouse=payload.get("warehouse_id")
            if requested_warehouse:
                warehouse=row("SELECT id FROM warehouses WHERE id=:id AND active=1",
                              {"id":requested_warehouse})
            if not warehouse:
                warehouse=row("SELECT id FROM warehouses WHERE code='WH-001' AND active=1 LIMIT 1")
            if not warehouse:
                warehouse=row("SELECT id FROM warehouses WHERE active=1 ORDER BY id LIMIT 1")
            if not warehouse:
                raise ValueError("لا يوجد مستودع نشط لتسجيل الرصيد الافتتاحي.")
        execute("""INSERT INTO inventory(code,sku,name,description,unit,quantity,unit_cost,cost,
          reorder_level,active)
          VALUES(:code,:code,:name,:description,:unit,0,:cost,:cost,:reorder,1)""",
          {"code":code,"name":name,"description":payload.get("description",""),
           "unit":normalize_item_unit(payload.get("unit")),
           "cost":cost,
           "reorder":float(payload.get("reorder_level") or 0)})
        new_id=row("SELECT id FROM inventory WHERE code=:code",{"code":code})["id"]
        if opening_quantity > 0:
            record_inventory_movement(
                date.today().isoformat(),"رصيد افتتاحي",new_id,warehouse["id"],
                opening_quantity,cost,reference_type="OPENING_BALANCE",
                reference_id=new_id,reference_no=code,
                notes="رصيد افتتاحي من الإنشاء السريع"
            )
        return "items",smart_entity_row("items",new_id)"""

source = replace_once(source, item_old, item_new, "quick item opening balance")

posting_old = """    a = require_accounts(["sales_account_id","vat_output_account_id"])
    customer_account_id = inv.get("customer_linked_account_id")
    if not customer_account_id:
        raise ValueError("العميل غير مربوط بحساب في دليل الحسابات. افتح بطاقة العميل وحدد حساب الذمم المدينة.")"""
posting_new = """    a = require_accounts(["customer_account_id","sales_account_id","vat_output_account_id"])
    customer_account_id = inv.get("customer_linked_account_id") or a["customer_account_id"]"""
source = replace_once(source, posting_old, posting_new, "invoice customer control account fallback")

if source == original:
    raise RuntimeError("No changes were produced.")

APP.write_text(source, encoding="utf-8")
print("Applied code changes to app.py")

assert 'SELECT COALESCE(SUM(amount),0) s FROM expenses' in source
assert 'record.get(field) if field in ("phone", "email")' in source
quick_block = source.split('if entity=="item":', 1)[1].split(
    'raise ValueError("نوع السجل غير مدعوم.")', 1
)[0]
assert '"رصيد افتتاحي"' in quick_block
assert 'customer_linked_account_id") or a["customer_account_id"]' in source

run(sys.executable, "-m", "py_compile", "app.py")
run("git", "diff", "--check")
print("Syntax check: PASS")
print("git diff --check: PASS")

pytest_probe = run(sys.executable, "-c", "import pytest", check=False)
if pytest_probe.returncode == 0:
    tests = run(sys.executable, "-m", "pytest", "-q", check=False)
    print(tests.stdout)
    if tests.returncode != 0:
        print(tests.stderr, file=sys.stderr)
        raise SystemExit("Existing automated tests failed; changes were NOT committed or pushed.")
    print("pytest: PASS")
else:
    print("pytest not installed; skipped automated pytest suite.")

print("\n--- git diff --stat ---")
print(run("git", "diff", "--stat").stdout)

run("git", "add", "app.py")
staged = run("git", "diff", "--cached", "--quiet", check=False)
if staged.returncode == 0:
    print("No staged changes; nothing to commit.")
else:
    run("git", "commit", "-m",
        "fix(staging): dashboard, contact persistence, opening stock and invoice posting")
    run("git", "push", "origin", EXPECTED_BRANCH)
    print(f"Pushed fixes to {EXPECTED_BRANCH}")

print("\nDONE. main and Production were not touched.")
