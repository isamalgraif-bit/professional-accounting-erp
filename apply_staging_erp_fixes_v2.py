#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
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
        f"SAFETY STOP: current branch is {branch!r}, expected {EXPECTED_BRANCH!r}. "
        "No files were changed."
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
    "generic contact preservation",
)

old_item_block = (
'    if entity=="item":\n'
'        code=next_inventory_sku()\n'
'        existing=row("SELECT id FROM inventory WHERE code=:code OR name=:name",\n'
'                     {"code":code,"name":name})\n'
'        if existing:\n'
'            raise ValueError("المادة موجودة مسبقًا.")\n'
'        execute("""INSERT INTO inventory(code,sku,name,description,unit,quantity,unit_cost,cost,\n'
'          reorder_level,active)\n'
'          VALUES(:code,:code,:name,:description,:unit,:quantity,:cost,:cost,:reorder,1)""",\n'
'          {"code":code,"name":name,"description":payload.get("description",""),\n'
'           "unit":normalize_item_unit(payload.get("unit")),\n'
'           "quantity":float(payload.get("quantity") or 0),\n'
'           "cost":float(payload.get("unit_cost") or 0),\n'
'           "reorder":float(payload.get("reorder_level") or 0)})\n'
'        new_id=row("SELECT id FROM inventory WHERE code=:code",{"code":code})["id"]\n'
'        return "items",smart_entity_row("items",new_id)'
)

new_item_block = (
'    if entity=="item":\n'
'        code=next_inventory_sku()\n'
'        existing=row("SELECT id FROM inventory WHERE code=:code OR name=:name",\n'
'                     {"code":code,"name":name})\n'
'        if existing:\n'
'            raise ValueError("المادة موجودة مسبقًا.")\n'
'        opening_quantity=round(float(payload.get("quantity") or 0),3)\n'
'        cost=round(float(payload.get("unit_cost") or 0),4)\n'
'        if opening_quantity < 0:\n'
'            raise ValueError("الرصيد الافتتاحي لا يمكن أن يكون سالبًا.")\n'
'        warehouse=None\n'
'        if opening_quantity > 0:\n'
'            requested_warehouse=payload.get("warehouse_id")\n'
'            if requested_warehouse:\n'
'                warehouse=row("SELECT id FROM warehouses WHERE id=:id AND active=1",\n'
'                              {"id":requested_warehouse})\n'
'            if not warehouse:\n'
'                warehouse=row("SELECT id FROM warehouses WHERE code=\'WH-001\' AND active=1 LIMIT 1")\n'
'            if not warehouse:\n'
'                warehouse=row("SELECT id FROM warehouses WHERE active=1 ORDER BY id LIMIT 1")\n'
'            if not warehouse:\n'
'                raise ValueError("لا يوجد مستودع نشط لتسجيل الرصيد الافتتاحي.")\n'
'        execute("""INSERT INTO inventory(code,sku,name,description,unit,quantity,unit_cost,cost,\n'
'          reorder_level,active)\n'
'          VALUES(:code,:code,:name,:description,:unit,0,:cost,:cost,:reorder,1)""",\n'
'          {"code":code,"name":name,"description":payload.get("description",""),\n'
'           "unit":normalize_item_unit(payload.get("unit")),\n'
'           "cost":cost,\n'
'           "reorder":float(payload.get("reorder_level") or 0)})\n'
'        new_id=row("SELECT id FROM inventory WHERE code=:code",{"code":code})["id"]\n'
'        if opening_quantity > 0:\n'
'            record_inventory_movement(\n'
'                datetime.now().date().isoformat(),"رصيد افتتاحي",new_id,warehouse["id"],\n'
'                opening_quantity,cost,reference_type="OPENING_BALANCE",\n'
'                reference_id=new_id,reference_no=code,\n'
'                notes="رصيد افتتاحي من الإنشاء السريع"\n'
'            )\n'
'        return "items",smart_entity_row("items",new_id)'
)

source = replace_once(source, old_item_block, new_item_block, "quick item opening balance")

old_posting = (
'    a = require_accounts(["sales_account_id","vat_output_account_id"])\n'
'    customer_account_id = inv.get("customer_linked_account_id")\n'
'    if not customer_account_id:\n'
'        raise ValueError("العميل غير مربوط بحساب في دليل الحسابات. افتح بطاقة العميل وحدد حساب الذمم المدينة.")'
)
new_posting = (
'    a = require_accounts(["customer_account_id","sales_account_id","vat_output_account_id"])\n'
'    customer_account_id = inv.get("customer_linked_account_id") or a["customer_account_id"]'
)

source = replace_once(source, old_posting, new_posting, "invoice posting account fallback")

if source == original:
    raise RuntimeError("No changes were produced.")

APP.write_text(source, encoding="utf-8")
print("Applied changes to app.py")

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
        raise SystemExit("Automated tests failed. Changes were NOT committed or pushed.")
    print("pytest: PASS")
else:
    print("pytest not installed; skipped.")

print("\n--- git diff --stat ---")
print(run("git", "diff", "--stat").stdout)

run("git", "add", "app.py")
staged = run("git", "diff", "--cached", "--quiet", check=False)
if staged.returncode == 0:
    print("No staged changes; nothing to commit.")
else:
    run(
        "git", "commit", "-m",
        "fix(staging): dashboard, contact persistence, opening stock and invoice posting"
    )
    run("git", "push", "origin", EXPECTED_BRANCH)
    print(f"Pushed fixes to {EXPECTED_BRANCH}")

print("\nDONE. main and Production were not touched.")
