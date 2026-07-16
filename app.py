from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import text
from functools import wraps
from datetime import datetime, timedelta
import os
import csv
import io
import base64
import uuid
from decimal import Decimal
import qrcode
from openpyxl import Workbook

APP_VERSION = "9.0.0"

app = Flask(__name__, template_folder=".")
app.secret_key = os.environ.get("SECRET_KEY", "development-only-change-me")

database_url = os.environ.get("DATABASE_URL", "sqlite:///erp.db")

if database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql+psycopg://", 1)
elif database_url.startswith("postgresql://") and not database_url.startswith("postgresql+psycopg://"):
    database_url = database_url.replace("postgresql://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True,
    "pool_recycle": 300,
}

db = SQLAlchemy(app)

def rows(query, params=None):
    result = db.session.execute(text(query), params or {})
    return result.mappings().all()

def row(query, params=None):
    result = db.session.execute(text(query), params or {})
    return result.mappings().first()

def execute(query, params=None):
    db.session.execute(text(query), params or {})
    db.session.commit()

def audit(action, entity, details=""):
    try:
        db.session.execute(
            text("""INSERT INTO audit_logs(user_id, username, action, entity, details, created_at)
                    VALUES(:user_id, :username, :action, :entity, :details, :created_at)"""),
            {
                "user_id": session.get("user_id"),
                "username": session.get("username", "system"),
                "action": action,
                "entity": entity,
                "details": details,
                "created_at": datetime.now(),
            },
        )
        db.session.commit()
    except Exception:
        db.session.rollback()


def next_invoice_number(invoice_date_value):
    """Create a yearly sequential invoice number such as INV-2026-000001."""
    if isinstance(invoice_date_value, str):
        invoice_year = datetime.strptime(invoice_date_value, "%Y-%m-%d").year
    else:
        invoice_year = invoice_date_value.year

    result = db.session.execute(
        text("""
            INSERT INTO invoice_sequences(sequence_year, last_number)
            VALUES(:year, 1)
            ON CONFLICT (sequence_year)
            DO UPDATE SET last_number = invoice_sequences.last_number + 1
            RETURNING last_number
        """),
        {"year": invoice_year},
    )
    sequence_number = result.scalar_one()
    db.session.commit()
    return f"INV-{invoice_year}-{sequence_number:06d}"



def next_journal_number(journal_date_value):
    if isinstance(journal_date_value, str):
        journal_year = datetime.strptime(journal_date_value, "%Y-%m-%d").year
    else:
        journal_year = journal_date_value.year

    result = db.session.execute(
        text("""
            INSERT INTO journal_sequences(sequence_year, last_number)
            VALUES(:year, 1)
            ON CONFLICT (sequence_year)
            DO UPDATE SET last_number = journal_sequences.last_number + 1
            RETURNING last_number
        """),
        {"year": journal_year},
    )
    number = result.scalar_one()
    db.session.commit()
    return f"JV-{journal_year}-{number:06d}"



def next_batch_number(batch_date_value):
    if isinstance(batch_date_value, str):
        batch_year = datetime.strptime(batch_date_value, "%Y-%m-%d").year
    else:
        batch_year = batch_date_value.year
    latest = db.session.execute(
        text("""SELECT COUNT(*) FROM journal_batches
                WHERE EXTRACT(YEAR FROM batch_date)=:year"""),
        {"year": batch_year}
    ).scalar() or 0
    return f"JB-{batch_year}-{latest + 1:06d}"



def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

def get_default_accounts():
    return row("""SELECT customer_account_id,supplier_account_id,sales_account_id,
                         purchases_account_id,vat_input_account_id,vat_output_account_id,
                         cash_account_id,bank_account_id,inventory_account_id,
                         cost_of_sales_account_id,retained_earnings_account_id
                  FROM settings WHERE id=1""")

def require_accounts(names):
    acc = get_default_accounts()
    missing = [n for n in names if not acc or not acc.get(n)]
    if missing:
        raise ValueError("أكمل ربط الحسابات الافتراضية من الإعدادات.")
    return acc

def ensure_open_period(document_date):
    p = row("""SELECT id,name,status FROM fiscal_periods
               WHERE :d BETWEEN start_date AND end_date
               ORDER BY start_date DESC LIMIT 1""", {"d":document_date})
    if not p:
        raise ValueError("لا توجد فترة مالية تغطي تاريخ المستند.")
    if p["status"] != "مفتوحة":
        raise ValueError(f"الفترة المالية {p['name']} مغلقة.")
    return p

def create_system_journal(journal_date, description, reference, source_type, source_id, lines):
    ensure_open_period(journal_date)
    td = round(sum(float(x.get("debit",0)) for x in lines),2)
    tc = round(sum(float(x.get("credit",0)) for x in lines),2)
    if td <= 0 or td != tc:
        raise ValueError("القيد التلقائي غير متوازن.")
    journal_no = next_journal_number(journal_date)
    db.session.execute(text("""INSERT INTO journal_entries(
        journal_no,journal_date,reference,description,status,total_debit,total_credit,
        created_by,created_at,source_type,source_id,posted_at)
        VALUES(:no,:dt,:ref,:des,'مرحّل',:td,:tc,:uid,:created,:stype,:sid,:posted)"""),
        {"no":journal_no,"dt":journal_date,"ref":reference,"des":description,
         "td":td,"tc":tc,"uid":session.get("user_id"),"created":datetime.now(),
         "stype":source_type,"sid":source_id,"posted":datetime.now()})
    journal_id = db.session.execute(text("SELECT id FROM journal_entries WHERE journal_no=:n"),
                                    {"n":journal_no}).scalar_one()
    for x in lines:
        db.session.execute(text("""INSERT INTO journal_entry_lines(
            journal_id,account_id,debit,credit,taxable,tax_direction,supplier_id,
            customer_id,party_type,tax_number,invoice_number,invoice_date,
            line_description,cost_center_id)
            VALUES(:jid,:aid,:debit,:credit,:taxable,:direction,:sid,:cid,:ptype,
                   :tax,:inv,:invdt,:des,:cc)"""),
            {"jid":journal_id,"aid":x["account_id"],"debit":round(float(x.get("debit",0)),2),
             "credit":round(float(x.get("credit",0)),2),"taxable":1 if x.get("taxable") else 0,
             "direction":x.get("tax_direction","غير مطبق"),"sid":x.get("supplier_id"),
             "cid":x.get("customer_id"),"ptype":x.get("party_type",""),
             "tax":x.get("tax_number",""),"inv":x.get("invoice_number",""),
             "invdt":x.get("invoice_date"),"des":x.get("line_description",""),
             "cc":x.get("cost_center_id")})
    db.session.commit()
    audit("POST",source_type,f"ترحيل تلقائي بالقيد {journal_no}")
    return journal_id

def post_invoice_to_ledger(invoice_id):
    inv = row("""SELECT i.*,c.vat_number customer_vat FROM invoices i
                 JOIN customers c ON c.id=i.customer_id WHERE i.id=:id""",{"id":invoice_id})
    if not inv: raise ValueError("الفاتورة غير موجودة.")
    if inv.get("journal_id"): return inv["journal_id"]
    a = require_accounts(["customer_account_id","sales_account_id","vat_output_account_id"])
    lines = [
      {"account_id":a["customer_account_id"],"debit":inv["total"],"customer_id":inv["customer_id"],
       "party_type":"عميل","tax_number":inv["customer_vat"] or "","invoice_number":inv["invoice_no"],
       "invoice_date":inv["invoice_date"],"line_description":"إجمالي فاتورة المبيعات"},
      {"account_id":a["sales_account_id"],"credit":inv["subtotal"],"customer_id":inv["customer_id"],
       "party_type":"عميل","tax_number":inv["customer_vat"] or "","invoice_number":inv["invoice_no"],
       "invoice_date":inv["invoice_date"],"line_description":"إيرادات المبيعات"},
    ]
    if float(inv["vat"] or 0):
        lines.append({"account_id":a["vat_output_account_id"],"credit":inv["vat"],"taxable":1,
          "tax_direction":"مخرجات","customer_id":inv["customer_id"],"party_type":"عميل",
          "tax_number":inv["customer_vat"] or "","invoice_number":inv["invoice_no"],
          "invoice_date":inv["invoice_date"],"line_description":"ضريبة المخرجات"})
    cost_total = sales_invoice_cost(invoice_id)
    if cost_total > 0:
        stock_accounts = require_accounts(["inventory_account_id","cost_of_sales_account_id"])
        lines.extend([
            {"account_id":stock_accounts["cost_of_sales_account_id"],"debit":cost_total,
             "invoice_number":inv["invoice_no"],"invoice_date":inv["invoice_date"],
             "line_description":"تكلفة البضاعة المباعة"},
            {"account_id":stock_accounts["inventory_account_id"],"credit":cost_total,
             "invoice_number":inv["invoice_no"],"invoice_date":inv["invoice_date"],
             "line_description":"إخراج المخزون المباع"},
        ])
    jid = create_system_journal(inv["invoice_date"],f"فاتورة مبيعات {inv['invoice_no']}",
                                inv["invoice_no"],"INVOICE",invoice_id,lines)
    execute("UPDATE invoices SET journal_id=:j,posting_status='مرحّل' WHERE id=:id",
            {"j":jid,"id":invoice_id})
    return jid

def post_expense_to_ledger(expense_id):
    e = row("""SELECT e.*,s.vat_number supplier_vat FROM expenses e
               LEFT JOIN suppliers s ON s.id=e.supplier_id WHERE e.id=:id""",{"id":expense_id})
    if not e: raise ValueError("المصروف غير موجود.")
    if e.get("journal_id"): return e["journal_id"]
    if not e["expense_account_id"] or not e["payment_account_id"]:
        raise ValueError("حدد حساب المصروف وحساب السداد.")
    a = require_accounts(["vat_input_account_id"] if float(e["vat"] or 0) else [])
    common = {"supplier_id":e["supplier_id"],"party_type":"مورد" if e["supplier_id"] else "",
              "tax_number":e["supplier_vat"] or "","invoice_number":e["invoice_number"] or "",
              "invoice_date":e["invoice_date"],"cost_center_id":e["cost_center_id"]}
    lines = [{"account_id":e["expense_account_id"],"debit":e["amount"],
              "line_description":e["description"] or e["category"],**common}]
    if float(e["vat"] or 0):
        lines.append({"account_id":a["vat_input_account_id"],"debit":e["vat"],"taxable":1,
                      "tax_direction":"مدخلات","line_description":"ضريبة المدخلات",**common})
    lines.append({"account_id":e["payment_account_id"],"credit":e["total"],
                  "line_description":"سداد المصروف",**common})
    jid = create_system_journal(e["expense_date"],f"مصروف {e['category']}",
                                e["invoice_number"] or f"EXP-{expense_id}",
                                "EXPENSE",expense_id,lines)
    execute("UPDATE expenses SET journal_id=:j,posting_status='مرحّل' WHERE id=:id",
            {"j":jid,"id":expense_id})
    return jid



ARABIC_TRANSLITERATION = {
    "ا":"a","أ":"a","إ":"i","آ":"aa","ب":"b","ت":"t","ث":"th","ج":"j","ح":"h",
    "خ":"kh","د":"d","ذ":"dh","ر":"r","ز":"z","س":"s","ش":"sh","ص":"s","ض":"d",
    "ط":"t","ظ":"z","ع":"a","غ":"gh","ف":"f","ق":"q","ك":"k","ل":"l","م":"m",
    "ن":"n","ه":"h","ة":"a","و":"w","ؤ":"o","ي":"y","ى":"a","ئ":"e","ء":"",
    "َ":"a","ُ":"u","ِ":"i","ّ":"","ْ":"","ً":"","ٌ":"","ٍ":"","ـ":""
}
BUSINESS_WORDS = {
    "شركة":"Company","الشركة":"Company","مؤسسة":"Establishment","المؤسسة":"Establishment",
    "محدودة":"Limited","المحدودة":"Limited","ذمم":"LLC","للمقاولات":"Contracting",
    "مقاولات":"Contracting","تجارة":"Trading","للتجارة":"Trading","صناعة":"Industries",
    "للصناعة":"Industries","خدمات":"Services","للخدمات":"Services","مجموعة":"Group",
    "مصنع":"Factory","مكتب":"Office","مركز":"Center","الدولية":"International",
    "العالمية":"International","العربية":"Arabian","السعودية":"Saudi","الخليج":"Gulf"
}

def transliterate_arabic_name(value):
    value = (value or "").strip()
    if not value:
        return ""
    words = value.split()
    result = []
    for word in words:
        normalized = word.strip("،,.-_/()")
        if normalized in BUSINESS_WORDS:
            result.append(BUSINESS_WORDS[normalized])
            continue
        latin = "".join(ARABIC_TRANSLITERATION.get(ch, ch) for ch in normalized)
        latin = re.sub(r"[^A-Za-z0-9]+", "", latin)
        if latin:
            result.append(latin[:1].upper() + latin[1:])
    # Improve common company suffix order.
    if result and result[0] == "Company":
        result = result[1:] + ["Company"]
    return " ".join(result)

@app.route("/api/transliterate", methods=["POST"])
@login_required
def transliterate_name_api():
    payload = request.get_json(silent=True) or {}
    return {"english_name": transliterate_arabic_name(payload.get("name",""))}



def next_treasury_number(voucher_type, voucher_date):
    prefix = {"قبض":"RV","صرف":"PV","تحويل":"TV"}.get(voucher_type,"TV")
    year = datetime.strptime(voucher_date,"%Y-%m-%d").year if isinstance(voucher_date,str) else voucher_date.year
    count = db.session.execute(text("""SELECT COUNT(*) FROM treasury_vouchers
        WHERE voucher_type=:t AND EXTRACT(YEAR FROM voucher_date)=:y"""),
        {"t":voucher_type,"y":year}).scalar() or 0
    return f"{prefix}-{year}-{count+1:06d}"

def post_treasury_voucher(voucher_id):
    v = row("""SELECT tv.*,c.name customer_name,c.vat_number customer_vat,
                     s.name supplier_name,s.vat_number supplier_vat
              FROM treasury_vouchers tv
              LEFT JOIN customers c ON c.id=tv.customer_id
              LEFT JOIN suppliers s ON s.id=tv.supplier_id
              WHERE tv.id=:id""", {"id":voucher_id})
    if not v:
        raise ValueError("السند غير موجود.")
    if v.get("journal_id"):
        return v["journal_id"]
    ensure_open_period(v["voucher_date"])

    common = {
        "customer_id":v["customer_id"],"supplier_id":v["supplier_id"],
        "party_type":v["party_type"],"tax_number":v["customer_vat"] or v["supplier_vat"] or "",
        "line_description":v["description"] or v["voucher_type"],
        "cost_center_id":v["cost_center_id"]
    }
    amount = float(v["amount"])
    if v["voucher_type"] == "قبض":
        lines = [
            {"account_id":v["cash_bank_account_id"],"debit":amount,**common},
            {"account_id":v["counter_account_id"],"credit":amount,**common},
        ]
    elif v["voucher_type"] == "صرف":
        lines = [
            {"account_id":v["counter_account_id"],"debit":amount,**common},
            {"account_id":v["cash_bank_account_id"],"credit":amount,**common},
        ]
    else:
        lines = [
            {"account_id":v["counter_account_id"],"debit":amount,**common},
            {"account_id":v["cash_bank_account_id"],"credit":amount,**common},
        ]

    jid = create_system_journal(
        v["voucher_date"], f"سند {v['voucher_type']} {v['voucher_no']}",
        v["reference"] or v["voucher_no"], "TREASURY", voucher_id, lines
    )
    execute("""UPDATE treasury_vouchers
               SET journal_id=:j,posting_status='مرحّل',status='معتمد'
               WHERE id=:id""", {"j":jid,"id":voucher_id})
    return jid



def next_document_number(table_name, date_field, prefix, document_date):
    year = datetime.strptime(document_date,"%Y-%m-%d").year if isinstance(document_date,str) else document_date.year
    allowed = {
        "purchase_requisitions": "requisition_date",
        "purchase_orders": "po_date",
        "goods_receipts": "grn_date",
        "supplier_invoices": "invoice_date",
    }
    if table_name not in allowed or allowed[table_name] != date_field:
        raise ValueError("Invalid document sequence.")
    count = db.session.execute(
        text(f"SELECT COUNT(*) FROM {table_name} WHERE EXTRACT(YEAR FROM {date_field})=:y"),
        {"y":year}
    ).scalar() or 0
    return f"{prefix}-{year}-{count+1:06d}"

def get_required_approver(document_type, amount):
    rule = row("""SELECT approver_role FROM approval_rules
                  WHERE document_type=:t AND active=1
                    AND :amount >= min_amount
                    AND (max_amount IS NULL OR :amount <= max_amount)
                  ORDER BY min_amount DESC LIMIT 1""",
               {"t":document_type,"amount":amount})
    return rule["approver_role"] if rule else "المدير المالي"

def post_supplier_invoice(invoice_id):
    inv = row("""SELECT si.*,s.vat_number supplier_vat
                 FROM supplier_invoices si
                 JOIN suppliers s ON s.id=si.supplier_id
                 WHERE si.id=:id""", {"id":invoice_id})
    if not inv:
        raise ValueError("فاتورة المورد غير موجودة.")
    if inv.get("journal_id"):
        return inv["journal_id"]
    acc = require_accounts(["supplier_account_id","vat_input_account_id"])
    common = {
        "supplier_id":inv["supplier_id"],"party_type":"مورد",
        "tax_number":inv["supplier_vat"] or "",
        "invoice_number":inv["supplier_invoice_no"],
        "invoice_date":inv["invoice_date"],
        "cost_center_id":inv["cost_center_id"],
    }
    lines = [
        {"account_id":inv["expense_or_inventory_account_id"],"debit":inv["subtotal"],
         "line_description":"فاتورة مورد - قيمة قبل الضريبة",**common}
    ]
    if float(inv["vat"] or 0) > 0:
        lines.append({"account_id":acc["vat_input_account_id"],"debit":inv["vat"],
                      "taxable":1,"tax_direction":"مدخلات",
                      "line_description":"ضريبة القيمة المضافة - مدخلات",**common})
    lines.append({"account_id":acc["supplier_account_id"],"credit":inv["total"],
                  "line_description":"ذمم المورد",**common})
    jid = create_system_journal(
        inv["invoice_date"], f"فاتورة مورد {inv['supplier_invoice_no']}",
        inv["supplier_invoice_no"], "SUPPLIER_INVOICE", invoice_id, lines
    )
    execute("""UPDATE supplier_invoices
               SET journal_id=:j,posting_status='مرحّل',status='معتمدة'
               WHERE id=:id""", {"j":jid,"id":invoice_id})
    return jid



def next_inventory_movement_number(movement_date):
    year = datetime.strptime(movement_date,"%Y-%m-%d").year if isinstance(movement_date,str) else movement_date.year
    count = db.session.execute(text("""SELECT COUNT(*) FROM inventory_movements
        WHERE EXTRACT(YEAR FROM movement_date)=:y"""),{"y":year}).scalar() or 0
    return f"IM-{year}-{count+1:07d}"

def next_inventory_count_number(count_date):
    year = datetime.strptime(count_date,"%Y-%m-%d").year if isinstance(count_date,str) else count_date.year
    count = db.session.execute(text("""SELECT COUNT(*) FROM inventory_counts
        WHERE EXTRACT(YEAR FROM count_date)=:y"""),{"y":year}).scalar() or 0
    return f"IC-{year}-{count+1:06d}"

def warehouse_stock(item_id, warehouse_id):
    value = db.session.execute(text("""
        SELECT COALESCE(SUM(
          CASE
            WHEN movement_type IN ('رصيد افتتاحي','استلام','تسوية زيادة','تحويل وارد','مرتجع مبيعات') THEN quantity
            WHEN movement_type IN ('صرف','تسوية نقص','تحويل صادر','بيع','مرتجع مشتريات') THEN -quantity
            ELSE 0 END
        ),0)
        FROM inventory_movements
        WHERE item_id=:item AND warehouse_id=:warehouse
    """),{"item":item_id,"warehouse":warehouse_id}).scalar()
    return float(value or 0)

def item_total_stock(item_id):
    value = db.session.execute(text("""
        SELECT COALESCE(SUM(
          CASE
            WHEN movement_type IN ('رصيد افتتاحي','استلام','تسوية زيادة','تحويل وارد','مرتجع مبيعات') THEN quantity
            WHEN movement_type IN ('صرف','تسوية نقص','تحويل صادر','بيع','مرتجع مشتريات') THEN -quantity
            ELSE 0 END
        ),0)
        FROM inventory_movements WHERE item_id=:item
    """),{"item":item_id}).scalar()
    return float(value or 0)

def record_inventory_movement(movement_date,movement_type,item_id,warehouse_id,quantity,
                              unit_cost=0,destination_warehouse_id=None,reference_type="",
                              reference_id=None,reference_no="",batch_no="",serial_no="",
                              production_date=None,expiry_date=None,notes=""):
    quantity=round(float(quantity),3)
    unit_cost=round(float(unit_cost or 0),4)
    if quantity<=0:
        raise ValueError("الكمية يجب أن تكون أكبر من صفر.")
    if movement_type in ("صرف","بيع","تحويل صادر","مرتجع مشتريات","تسوية نقص"):
        available=warehouse_stock(item_id,warehouse_id)
        if quantity>available+0.0001:
            raise ValueError(f"الرصيد غير كافٍ. المتاح {available:.3f}")
    no=next_inventory_movement_number(movement_date)
    execute("""INSERT INTO inventory_movements(
      movement_no,movement_date,movement_type,item_id,warehouse_id,destination_warehouse_id,
      quantity,unit_cost,total_cost,reference_type,reference_id,reference_no,batch_no,
      serial_no,production_date,expiry_date,notes,created_by,created_at)
      VALUES(:no,:dt,:type,:item,:warehouse,:dest,:qty,:cost,:total,:rtype,:rid,:rno,
      :batch,:serial,:prod,:expiry,:notes,:uid,:created)""",
      {"no":no,"dt":movement_date,"type":movement_type,"item":item_id,
       "warehouse":warehouse_id,"dest":destination_warehouse_id,"qty":quantity,
       "cost":unit_cost,"total":round(quantity*unit_cost,2),"rtype":reference_type,
       "rid":reference_id,"rno":reference_no,"batch":batch_no,"serial":serial_no,
       "prod":production_date,"expiry":expiry_date,"notes":notes,
       "uid":session.get("user_id"),"created":datetime.now()})
    # Moving-average cost on incoming movements.
    if movement_type in ("رصيد افتتاحي","استلام","تسوية زيادة","تحويل وارد","مرتجع مبيعات"):
        item=row("SELECT cost FROM inventory WHERE id=:id",{"id":item_id})
        previous_qty=max(item_total_stock(item_id)-quantity,0)
        old_cost=float(item["cost"] or 0)
        new_cost=((previous_qty*old_cost)+(quantity*unit_cost))/(previous_qty+quantity) if previous_qty+quantity else unit_cost
        execute("UPDATE inventory SET quantity=:q,cost=:c WHERE id=:id",
                {"q":item_total_stock(item_id),"c":round(new_cost,4),"id":item_id})
    else:
        execute("UPDATE inventory SET quantity=:q WHERE id=:id",
                {"q":item_total_stock(item_id),"id":item_id})
    return no



def next_sales_no(table_name, date_field, prefix, doc_date):
    allowed={"sales_quotations":"quotation_date","sales_orders":"order_date","sales_deliveries":"delivery_date"}
    if allowed.get(table_name)!=date_field: raise ValueError("Invalid sequence")
    y=datetime.strptime(doc_date,"%Y-%m-%d").year if isinstance(doc_date,str) else doc_date.year
    n=db.session.execute(text(f"SELECT COUNT(*) FROM {table_name} WHERE EXTRACT(YEAR FROM {date_field})=:y"),{"y":y}).scalar() or 0
    return f"{prefix}-{y}-{n+1:06d}"

def parse_sales_lines(form):
    ids=form.getlist("item_id[]"); qtys=form.getlist("quantity[]"); prices=form.getlist("unit_price[]")
    discounts=form.getlist("discount_rate[]"); vats=form.getlist("vat_rate[]")
    lines=[]; subtotal=discount_total=vat_total=total=0
    for i,item_id in enumerate(ids):
        if not item_id: continue
        item=row("SELECT id,sku,name,unit,sale_price FROM inventory WHERE id=:id",{"id":item_id})
        qty=float(qtys[i] or 0); price=float(prices[i] or item["sale_price"] or 0)
        dr=float(discounts[i] or 0); vr=float(vats[i] or 0)
        if qty<=0: raise ValueError("الكمية يجب أن تكون أكبر من صفر")
        base=round(qty*price,2); disc=round(base*dr/100,2); taxable=base-disc
        tax=round(taxable*vr/100,2); line_total=round(taxable+tax,2)
        lines.append({"item_id":item["id"],"item_name":item["name"],"quantity":qty,"unit":item["unit"],
                      "unit_price":price,"discount_rate":dr,"vat_rate":vr,
                      "line_subtotal":base,"line_discount":disc,"line_vat":tax,"line_total":line_total})
        subtotal+=base; discount_total+=disc; vat_total+=tax; total+=line_total
    if not lines: raise ValueError("أدخل صنفًا واحدًا على الأقل")
    return lines,round(subtotal,2),round(discount_total,2),round(vat_total,2),round(total,2)


def create_invoice_from_sales_delivery(delivery_id):
    delivery = row("""SELECT d.*,o.order_no,o.branch_id,o.cost_center_id,o.subtotal,o.discount,
                             o.vat,o.total
                      FROM sales_deliveries d
                      JOIN sales_orders o ON o.id=d.order_id
                      WHERE d.id=:id""", {"id":delivery_id})
    if not delivery:
        raise ValueError("إذن التسليم غير موجود.")
    if delivery.get("invoice_id"):
        return delivery["invoice_id"]

    delivery_items = rows("""SELECT di.*,oi.item_name,oi.unit,oi.unit_price,oi.vat_rate,
                                    oi.discount_rate
                             FROM sales_delivery_items di
                             JOIN sales_order_items oi ON oi.id=di.order_item_id
                             WHERE di.delivery_id=:id ORDER BY di.id""", {"id":delivery_id})
    if not delivery_items:
        raise ValueError("لا توجد بنود في إذن التسليم.")

    subtotal = vat_total = grand_total = 0.0
    prepared = []
    for x in delivery_items:
        base = round(float(x["quantity"]) * float(x["unit_price"]), 2)
        discount = round(base * float(x["discount_rate"] or 0) / 100, 2)
        taxable = round(base - discount, 2)
        tax = round(taxable * float(x["vat_rate"] or 0) / 100, 2)
        total = round(taxable + tax, 2)
        prepared.append({
            "item_id": x["item_id"], "item_name": x["item_name"], "quantity": x["quantity"],
            "unit": x["unit"], "unit_price": x["unit_price"], "vat_rate": x["vat_rate"],
            "line_subtotal": taxable, "line_vat": tax, "line_total": total
        })
        subtotal += taxable
        vat_total += tax
        grand_total += total

    invoice_no = next_invoice_number(delivery["delivery_date"])
    invoice_uuid = str(uuid.uuid4())
    execute("""INSERT INTO invoices(
        invoice_no,invoice_uuid,customer_id,invoice_date,subtotal,vat,total,status,
        branch_id,notes,created_at,sales_order_id,delivery_id,cost_center_id,warehouse_id
        ) VALUES(
        :no,:uuid,:customer,:dt,:sub,:vat,:total,'مسودة',:branch,:notes,:created,
        :order_id,:delivery_id,:cc,:warehouse
        )""", {
        "no":invoice_no,"uuid":invoice_uuid,"customer":delivery["customer_id"],
        "dt":delivery["delivery_date"],"sub":round(subtotal,2),"vat":round(vat_total,2),
        "total":round(grand_total,2),"branch":delivery["branch_id"],
        "notes":f"فاتورة ناتجة عن إذن التسليم {delivery['delivery_no']}",
        "created":datetime.now(),"order_id":delivery["order_id"],"delivery_id":delivery_id,
        "cc":delivery["cost_center_id"],"warehouse":delivery["warehouse_id"]
    })
    order_meta = row("""SELECT payment_terms,sales_person FROM sales_orders WHERE id=:id""",
                     {"id":delivery["order_id"]})
    due_date = delivery["delivery_date"]
    if order_meta and order_meta.get("payment_terms"):
        match = re.search(r"(\d+)", order_meta["payment_terms"])
        if match:
            due_date = delivery["delivery_date"] + timedelta(days=int(match.group(1)))
    execute("""UPDATE invoices SET due_date=:due,payment_terms=:terms,
               sales_person=:sales_person,payment_method='آجل'
               WHERE id=:id""",
            {"due":due_date,
             "terms":order_meta["payment_terms"] if order_meta else "",
             "sales_person":order_meta["sales_person"] if order_meta else "",
             "id":invoice_id})
    invoice_id = row("SELECT id FROM invoices WHERE invoice_no=:no",{"no":invoice_no})["id"]

    for x in prepared:
        execute("""INSERT INTO invoice_items(
            invoice_id,item_id,item_name,quantity,unit,unit_price,vat_rate,
            line_subtotal,line_vat,line_total
            ) VALUES(
            :invoice_id,:item_id,:item_name,:quantity,:unit,:unit_price,:vat_rate,
            :line_subtotal,:line_vat,:line_total
            )""", {"invoice_id":invoice_id, **x})

    execute("UPDATE sales_deliveries SET invoice_id=:invoice WHERE id=:id",
            {"invoice":invoice_id,"id":delivery_id})
    audit("CREATE","INVOICE",f"إنشاء فاتورة {invoice_no} من إذن التسليم {delivery['delivery_no']}")
    return invoice_id

def sales_invoice_cost(invoice_id):
    value = db.session.execute(text("""
        SELECT COALESCE(SUM(di.quantity * di.unit_cost),0)
        FROM sales_delivery_items di
        JOIN sales_deliveries d ON d.id=di.delivery_id
        JOIN invoices i ON i.delivery_id=d.id
        WHERE i.id=:invoice_id
    """), {"invoice_id":invoice_id}).scalar()
    if value:
        return float(value)
    value = db.session.execute(text("""
        SELECT COALESCE(SUM(ii.quantity * COALESCE(inv.cost,0)),0)
        FROM invoice_items ii LEFT JOIN inventory inv ON inv.id=ii.item_id
        WHERE ii.invoice_id=:invoice_id
    """), {"invoice_id":invoice_id}).scalar()
    return float(value or 0)

def post_sales_return_to_ledger(return_id):
    ret = row("""SELECT r.*,c.vat_number customer_vat
                 FROM sales_returns r JOIN customers c ON c.id=r.customer_id
                 WHERE r.id=:id""", {"id":return_id})
    if not ret:
        raise ValueError("مرتجع المبيعات غير موجود.")
    if ret.get("journal_id"):
        return ret["journal_id"]

    acc = require_accounts(["customer_account_id","sales_account_id","vat_output_account_id",
                            "inventory_account_id","cost_of_sales_account_id"])
    common = {
        "customer_id":ret["customer_id"],"party_type":"عميل",
        "tax_number":ret["customer_vat"] or "","invoice_number":ret["return_no"],
        "invoice_date":ret["return_date"],"line_description":"مرتجع مبيعات"
    }
    lines = [
        {"account_id":acc["sales_account_id"],"debit":ret["subtotal"],**common},
        {"account_id":acc["customer_account_id"],"credit":ret["total"],**common},
    ]
    if float(ret["vat"] or 0) > 0:
        lines.append({"account_id":acc["vat_output_account_id"],"debit":ret["vat"],
                      "taxable":1,"tax_direction":"مخرجات",**common})
    if float(ret["cost_total"] or 0) > 0:
        lines.extend([
            {"account_id":acc["inventory_account_id"],"debit":ret["cost_total"],**common},
            {"account_id":acc["cost_of_sales_account_id"],"credit":ret["cost_total"],**common},
        ])

    jid = create_system_journal(
        ret["return_date"],f"مرتجع مبيعات {ret['return_no']}",
        ret["return_no"],"SALES_RETURN",return_id,lines
    )
    execute("UPDATE sales_returns SET journal_id=:jid WHERE id=:id",
            {"jid":jid,"id":return_id})
    return jid



def invoice_paid_amount(invoice_id):
    value = db.session.execute(text("""
        SELECT COALESCE(SUM(allocated_amount),0)
        FROM invoice_payment_allocations WHERE invoice_id=:id
    """), {"id":invoice_id}).scalar()
    return round(float(value or 0),2)

def refresh_invoice_payment_status(invoice_id):
    inv = row("SELECT total FROM invoices WHERE id=:id",{"id":invoice_id})
    if not inv:
        return
    paid = invoice_paid_amount(invoice_id)
    total = round(float(inv["total"] or 0),2)
    if paid <= 0:
        status = "غير مسددة"
    elif paid + 0.005 >= total:
        status = "مسددة"
    else:
        status = "مسددة جزئيًا"
    execute("UPDATE invoices SET payment_status=:status WHERE id=:id",
            {"status":status,"id":invoice_id})

def allocate_customer_receipt(voucher_id, allocations):
    voucher = row("""SELECT * FROM treasury_vouchers
                     WHERE id=:id AND voucher_type='قبض' AND customer_id IS NOT NULL""",
                  {"id":voucher_id})
    if not voucher:
        raise ValueError("سند القبض غير صالح للربط بالفواتير.")
    allocation_total = round(sum(float(x["amount"]) for x in allocations),2)
    if allocation_total <= 0:
        raise ValueError("يجب تخصيص مبلغ لفاتورة واحدة على الأقل.")
    if allocation_total > round(float(voucher["amount"]),2) + 0.005:
        raise ValueError("إجمالي التخصيص أكبر من مبلغ سند القبض.")

    for allocation in allocations:
        invoice_id = int(allocation["invoice_id"])
        amount = round(float(allocation["amount"]),2)
        if amount <= 0:
            continue
        inv = row("""SELECT id,total,customer_id FROM invoices
                     WHERE id=:id AND customer_id=:customer""",
                  {"id":invoice_id,"customer":voucher["customer_id"]})
        if not inv:
            raise ValueError("إحدى الفواتير لا تخص العميل المحدد.")
        outstanding = round(float(inv["total"]) - invoice_paid_amount(invoice_id),2)
        if amount > outstanding + 0.005:
            raise ValueError(f"المبلغ المخصص أكبر من رصيد الفاتورة المتبقي {outstanding:.2f}.")
        execute("""INSERT INTO invoice_payment_allocations(
                    invoice_id,voucher_id,allocated_amount,allocation_date,created_by,created_at)
                   VALUES(:invoice,:voucher,:amount,:dt,:uid,:created)""",
                {"invoice":invoice_id,"voucher":voucher_id,"amount":amount,
                 "dt":voucher["voucher_date"],"uid":session.get("user_id"),
                 "created":datetime.now()})
        refresh_invoice_payment_status(invoice_id)



def sales_document_timeline(invoice_id=None, delivery_id=None, order_id=None, quotation_id=None):
    context = {
        "quotation": None, "order": None, "delivery": None,
        "invoice": None, "receipt_allocations": [], "return_docs": []
    }

    if invoice_id:
        context["invoice"] = row("""SELECT i.*,c.name customer_name
                                    FROM invoices i JOIN customers c ON c.id=i.customer_id
                                    WHERE i.id=:id""", {"id":invoice_id})
        if context["invoice"]:
            delivery_id = delivery_id or context["invoice"].get("delivery_id")
            order_id = order_id or context["invoice"].get("sales_order_id")
            context["receipt_allocations"] = rows("""SELECT a.*,tv.voucher_no,tv.voucher_date,
                tv.payment_method,tv.reference
                FROM invoice_payment_allocations a
                JOIN treasury_vouchers tv ON tv.id=a.voucher_id
                WHERE a.invoice_id=:id ORDER BY tv.voucher_date,tv.id""", {"id":invoice_id})
            context["return_docs"] = rows("""SELECT id,return_no,return_date,total,status
                                             FROM sales_returns WHERE invoice_id=:id
                                             ORDER BY return_date,id""", {"id":invoice_id})

    if delivery_id:
        context["delivery"] = row("""SELECT d.*,o.order_no
                                     FROM sales_deliveries d
                                     JOIN sales_orders o ON o.id=d.order_id
                                     WHERE d.id=:id""", {"id":delivery_id})
        if context["delivery"]:
            order_id = order_id or context["delivery"]["order_id"]

    if order_id:
        context["order"] = row("""SELECT o.*,q.quotation_no
                                  FROM sales_orders o
                                  LEFT JOIN sales_quotations q ON q.id=o.quotation_id
                                  WHERE o.id=:id""", {"id":order_id})
        if context["order"]:
            quotation_id = quotation_id or context["order"].get("quotation_id")

    if quotation_id:
        context["quotation"] = row("SELECT * FROM sales_quotations WHERE id=:id",
                                   {"id":quotation_id})

    return context



def financial_date_filters():
    return {
        "date_from": request.args.get("date_from","").strip(),
        "date_to": request.args.get("date_to","").strip(),
        "branch_id": request.args.get("branch_id","").strip(),
        "cost_center_id": request.args.get("cost_center_id","").strip(),
    }

def financial_where(filters, journal_alias="j", line_alias="l"):
    conditions=[f"{journal_alias}.status='مرحّل'"]
    params={}
    if filters.get("date_from"):
        conditions.append(f"{journal_alias}.journal_date>=:date_from")
        params["date_from"]=filters["date_from"]
    if filters.get("date_to"):
        conditions.append(f"{journal_alias}.journal_date<=:date_to")
        params["date_to"]=filters["date_to"]
    if filters.get("cost_center_id"):
        conditions.append(f"{line_alias}.cost_center_id=:cost_center_id")
        params["cost_center_id"]=filters["cost_center_id"]
    return " AND ".join(conditions),params

def account_signed_balance(account_type, debit, credit):
    debit=float(debit or 0); credit=float(credit or 0)
    if account_type in ("التزام","حقوق ملكية","إيراد"):
        return credit-debit
    return debit-credit

def financial_statement_data(filters):
    where,params=financial_where(filters)
    data=rows(f"""SELECT a.id,a.account_code,a.account_name_ar,a.account_name_en,
      a.account_type,a.statement_type,a.normal_balance,
      COALESCE(SUM(l.debit),0) total_debit,COALESCE(SUM(l.credit),0) total_credit
      FROM chart_of_accounts a
      LEFT JOIN journal_entry_lines l ON l.account_id=a.id
      LEFT JOIN journal_entries j ON j.id=l.journal_id AND {where}
      WHERE a.active=1
      GROUP BY a.id,a.account_code,a.account_name_ar,a.account_name_en,
               a.account_type,a.statement_type,a.normal_balance
      ORDER BY a.account_code""",params)
    result=[]
    for r in data:
        item=dict(r)
        item["balance"]=round(account_signed_balance(
            item["account_type"],item["total_debit"],item["total_credit"]),2)
        result.append(item)
    return result

def party_statement_rows(party_type, party_id, date_from="", date_to=""):
    field="customer_id" if party_type=="customer" else "supplier_id"
    conditions=[f"l.{field}=:party","j.status='مرحّل'"]
    params={"party":party_id}
    if date_from:
        conditions.append("j.journal_date>=:date_from");params["date_from"]=date_from
    if date_to:
        conditions.append("j.journal_date<=:date_to");params["date_to"]=date_to
    data=rows(f"""SELECT j.id journal_id,j.journal_no,j.journal_date,j.reference,
      j.description,l.invoice_number,l.invoice_date,l.line_description,
      l.debit,l.credit,a.account_code,a.account_name_ar
      FROM journal_entry_lines l
      JOIN journal_entries j ON j.id=l.journal_id
      JOIN chart_of_accounts a ON a.id=l.account_id
      WHERE {' AND '.join(conditions)}
      ORDER BY j.journal_date,j.id,l.id""",params)
    running=0
    output=[]
    for r in data:
        item=dict(r)
        running += float(item["debit"] or 0)-float(item["credit"] or 0)
        item["running_balance"]=round(running,2)
        output.append(item)
    return output



def next_asset_number():
    count = db.session.execute(text("SELECT COUNT(*) FROM fixed_assets")).scalar() or 0
    return f"FA-{count+1:06d}"

def next_asset_run_number(period_date):
    year = period_date.year if hasattr(period_date,"year") else datetime.strptime(period_date,"%Y-%m-%d").year
    count = db.session.execute(text("""SELECT COUNT(*) FROM asset_depreciation_runs
        WHERE EXTRACT(YEAR FROM period_date)=:y"""),{"y":year}).scalar() or 0
    return f"DEP-{year}-{count+1:05d}"

def calculate_monthly_depreciation(asset):
    cost=float(asset["purchase_cost"] or 0)
    residual=float(asset["residual_value"] or 0)
    life=int(asset["useful_life_months"] or 1)
    accumulated=float(asset["accumulated_depreciation"] or 0)
    depreciable=max(cost-residual,0)
    remaining=max(depreciable-accumulated,0)
    if remaining<=0:
        return 0
    if asset["depreciation_method"]=="القسط الثابت":
        return round(min(depreciable/life,remaining),2)
    return round(min(depreciable/life,remaining),2)

def post_depreciation_run(run_id):
    run=row("SELECT * FROM asset_depreciation_runs WHERE id=:id",{"id":run_id})
    if not run:
        raise ValueError("دفعة الإهلاك غير موجودة.")
    if run.get("journal_id"):
        return run["journal_id"]
    lines=rows("""SELECT dl.*,fa.name,ac.depreciation_expense_account_id,
      ac.accumulated_depr_account_id
      FROM asset_depreciation_lines dl
      JOIN fixed_assets fa ON fa.id=dl.asset_id
      JOIN asset_categories ac ON ac.id=fa.category_id
      WHERE dl.run_id=:id""",{"id":run_id})
    journal_lines=[]
    for x in lines:
        if not x["depreciation_expense_account_id"] or not x["accumulated_depr_account_id"]:
            raise ValueError(f"فئة الأصل {x['name']} لا تحتوي على حسابات الإهلاك.")
        journal_lines.extend([
            {"account_id":x["depreciation_expense_account_id"],"debit":x["depreciation_amount"],
             "line_description":f"إهلاك الأصل {x['name']}"},
            {"account_id":x["accumulated_depr_account_id"],"credit":x["depreciation_amount"],
             "line_description":f"مجمع إهلاك الأصل {x['name']}"},
        ])
    jid=create_system_journal(run["period_date"],f"إهلاك الأصول {run['run_no']}",
                              run["run_no"],"ASSET_DEPRECIATION",run_id,journal_lines)
    execute("UPDATE asset_depreciation_runs SET journal_id=:jid,status='مرحّل' WHERE id=:id",
            {"jid":jid,"id":run_id})
    return jid



def next_payroll_run_number(period_end):
    year = period_end.year if hasattr(period_end,"year") else datetime.strptime(period_end,"%Y-%m-%d").year
    count = db.session.execute(text("""SELECT COUNT(*) FROM payroll_runs
        WHERE EXTRACT(YEAR FROM period_end)=:y"""),{"y":year}).scalar() or 0
    return f"PAY-{year}-{count+1:05d}"

def payroll_configuration():
    config=row("SELECT * FROM payroll_settings WHERE id=1")
    if not config:
        execute("""INSERT INTO payroll_settings(id,working_days_per_month,
                   working_hours_per_day,overtime_multiplier)
                   VALUES(1,30,8,1.5)""")
        config=row("SELECT * FROM payroll_settings WHERE id=1")
    return config

def calculate_employee_payroll(employee, period_start, period_end):
    cfg=payroll_configuration()
    working_days=float(cfg["working_days_per_month"] or 30)
    hours_per_day=float(cfg["working_hours_per_day"] or 8)
    overtime_multiplier=float(cfg["overtime_multiplier"] or 1.5)

    attendance=row("""SELECT
      COALESCE(SUM(overtime_hours),0) overtime_hours,
      COALESCE(SUM(CASE WHEN status='غائب' THEN 1 ELSE 0 END),0) absence_days,
      COALESCE(SUM(absence_hours),0) absence_hours
      FROM attendance_records
      WHERE employee_id=:employee AND attendance_date BETWEEN :start AND :end""",
      {"employee":employee["id"],"start":period_start,"end":period_end})

    basic=float(employee["basic_salary"] or 0)
    allowances=(float(employee.get("allowances") or 0)
                +float(employee.get("housing_allowance") or 0)
                +float(employee.get("transport_allowance") or 0)
                +float(employee.get("other_allowance") or 0))
    hourly_rate=basic/working_days/hours_per_day if working_days and hours_per_day else 0
    overtime_hours=float(attendance["overtime_hours"] or 0)
    overtime_amount=round(overtime_hours*hourly_rate*overtime_multiplier,2)
    absence_days=float(attendance["absence_days"] or 0)
    absence_hours=float(attendance["absence_hours"] or 0)
    absence_deduction=round((absence_days*basic/working_days)+(absence_hours*hourly_rate),2)

    adjustments=rows("""SELECT adjustment_type,amount FROM employee_salary_adjustments
      WHERE employee_id=:employee AND active=1 AND adjustment_date<=:end
      AND (recurring=1 OR adjustment_date BETWEEN :start AND :end)""",
      {"employee":employee["id"],"start":period_start,"end":period_end})
    extra_earnings=sum(float(x["amount"]) for x in adjustments if x["adjustment_type"]=="استحقاق")
    other_deductions=sum(float(x["amount"]) for x in adjustments if x["adjustment_type"]=="استقطاع")

    gross=round(basic+allowances+overtime_amount+extra_earnings,2)
    deductions=round(absence_deduction+other_deductions,2)
    net=round(max(gross-deductions,0),2)
    return {
        "basic_salary":basic,"allowances":round(allowances+extra_earnings,2),
        "overtime_hours":overtime_hours,"overtime_amount":overtime_amount,
        "absence_days":absence_days,"absence_deduction":absence_deduction,
        "other_deductions":round(other_deductions,2),"gross_salary":gross,
        "total_deductions":deductions,"net_salary":net
    }

def post_payroll_run(run_id):
    run=row("SELECT * FROM payroll_runs WHERE id=:id",{"id":run_id})
    if not run:
        raise ValueError("مسير الرواتب غير موجود.")
    if run.get("journal_id"):
        return run["journal_id"]
    cfg=payroll_configuration()
    required=["salary_expense_account_id","payroll_payable_account_id","deduction_account_id"]
    missing=[x for x in required if not cfg.get(x)]
    if missing:
        raise ValueError("أكمل ربط حسابات الرواتب من إعدادات الرواتب.")

    lines=[]
    for x in rows("""SELECT pl.*,e.name FROM payroll_lines pl
                     JOIN employees e ON e.id=pl.employee_id
                     WHERE pl.run_id=:id""",{"id":run_id}):
        common={"cost_center_id":x["cost_center_id"],
                "line_description":f"راتب الموظف {x['name']}"}
        lines.append({"account_id":cfg["salary_expense_account_id"],
                      "debit":x["gross_salary"],**common})
        lines.append({"account_id":cfg["payroll_payable_account_id"],
                      "credit":x["net_salary"],**common})
        if float(x["total_deductions"] or 0)>0:
            lines.append({"account_id":cfg["deduction_account_id"],
                          "credit":x["total_deductions"],**common})
    jid=create_system_journal(run["payment_date"] or run["period_end"],
                              f"مسير الرواتب {run['run_no']}",run["run_no"],
                              "PAYROLL",run_id,lines)
    execute("""UPDATE payroll_runs SET journal_id=:jid,status='مرحّل'
               WHERE id=:id""",{"jid":jid,"id":run_id})
    return jid


def init_db():
    statements = [
        """CREATE TABLE IF NOT EXISTS users(
            id SERIAL PRIMARY KEY,
            username VARCHAR(100) UNIQUE NOT NULL,
            password VARCHAR(255) NOT NULL,
            role VARCHAR(50) NOT NULL DEFAULT 'admin'
        )""",
        """CREATE TABLE IF NOT EXISTS settings(
            id INTEGER PRIMARY KEY,
            company_name_ar VARCHAR(255) DEFAULT 'اسم الشركة',
            company_name_en VARCHAR(255) DEFAULT 'Company Name',
            vat_number VARCHAR(50) DEFAULT '',
            cr_number VARCHAR(50) DEFAULT '',
            currency VARCHAR(10) DEFAULT 'SAR',
            address TEXT DEFAULT '',
            phone VARCHAR(50) DEFAULT '',
            email VARCHAR(255) DEFAULT '',
            logo_url TEXT DEFAULT '',
            logo_data TEXT DEFAULT '',
            logo_mime VARCHAR(50) DEFAULT '',
            vat_rate NUMERIC(5,2) DEFAULT 15.00
        )""",
        """CREATE TABLE IF NOT EXISTS branches(
            id SERIAL PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
            city VARCHAR(150),
            active INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS customers(
            id SERIAL PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
            name_en VARCHAR(255) DEFAULT '',
            vat_number VARCHAR(50),
            phone VARCHAR(50),
            email VARCHAR(255),
            balance NUMERIC(18,2) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS suppliers(
            id SERIAL PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
            name_en VARCHAR(255) DEFAULT '',
            vat_number VARCHAR(50),
            phone VARCHAR(50),
            email VARCHAR(255),
            balance NUMERIC(18,2) DEFAULT 0
        )""",

        """CREATE TABLE IF NOT EXISTS invoice_sequences(
            sequence_year INTEGER PRIMARY KEY,
            last_number INTEGER NOT NULL DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS invoices(
            id SERIAL PRIMARY KEY,
            invoice_no VARCHAR(100) UNIQUE NOT NULL,
            invoice_uuid VARCHAR(64),
            customer_id INTEGER NOT NULL REFERENCES customers(id),
            invoice_date DATE NOT NULL,
            subtotal NUMERIC(18,2) NOT NULL,
            vat NUMERIC(18,2) NOT NULL,
            total NUMERIC(18,2) NOT NULL,
            status VARCHAR(50) NOT NULL DEFAULT 'مسودة',
            branch_id INTEGER REFERENCES branches(id),
            notes TEXT,
            created_at TIMESTAMP NOT NULL
        )""",

        """CREATE TABLE IF NOT EXISTS invoice_items(
            id SERIAL PRIMARY KEY,
            invoice_id INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE,
            item_name VARCHAR(255) NOT NULL,
            description TEXT DEFAULT '',
            quantity NUMERIC(18,3) NOT NULL DEFAULT 1,
            unit VARCHAR(50) DEFAULT 'وحدة',
            unit_price NUMERIC(18,2) NOT NULL DEFAULT 0,
            vat_rate NUMERIC(5,2) NOT NULL DEFAULT 15,
            line_subtotal NUMERIC(18,2) NOT NULL DEFAULT 0,
            line_vat NUMERIC(18,2) NOT NULL DEFAULT 0,
            line_total NUMERIC(18,2) NOT NULL DEFAULT 0
        )""",


        """CREATE TABLE IF NOT EXISTS chart_of_accounts(
            id SERIAL PRIMARY KEY,
            account_code VARCHAR(50) UNIQUE NOT NULL,
            account_name_ar VARCHAR(255) NOT NULL,
            account_name_en VARCHAR(255) DEFAULT '',
            account_type VARCHAR(50) NOT NULL,
            parent_id INTEGER REFERENCES chart_of_accounts(id),
            level INTEGER NOT NULL DEFAULT 1,
            accepts_entries INTEGER NOT NULL DEFAULT 1,
            normal_balance VARCHAR(20) DEFAULT 'مدين',
            statement_type VARCHAR(50) DEFAULT 'الميزانية العمومية',
            active INTEGER NOT NULL DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS cost_centers(
            id SERIAL PRIMARY KEY,
            code VARCHAR(50) UNIQUE NOT NULL,
            name VARCHAR(255) NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS journal_sequences(
            sequence_year INTEGER PRIMARY KEY,
            last_number INTEGER NOT NULL DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS journal_entries(
            id SERIAL PRIMARY KEY,
            journal_no VARCHAR(100) UNIQUE NOT NULL,
            journal_date DATE NOT NULL,
            reference VARCHAR(255) DEFAULT '',
            description TEXT DEFAULT '',
            status VARCHAR(50) NOT NULL DEFAULT 'مسودة',
            total_debit NUMERIC(18,2) NOT NULL DEFAULT 0,
            total_credit NUMERIC(18,2) NOT NULL DEFAULT 0,
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS journal_entry_lines(
            id SERIAL PRIMARY KEY,
            journal_id INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
            account_id INTEGER NOT NULL REFERENCES chart_of_accounts(id),
            debit NUMERIC(18,2) NOT NULL DEFAULT 0,
            credit NUMERIC(18,2) NOT NULL DEFAULT 0,
            taxable INTEGER NOT NULL DEFAULT 0,
            tax_direction VARCHAR(50) DEFAULT 'غير مطبق',
            supplier_id INTEGER REFERENCES suppliers(id),
            customer_id INTEGER REFERENCES customers(id),
            party_type VARCHAR(20) DEFAULT '',
            tax_number VARCHAR(50) DEFAULT '',
            invoice_number VARCHAR(100) DEFAULT '',
            line_description TEXT DEFAULT '',
            cost_center_id INTEGER REFERENCES cost_centers(id)
        )""",

        """CREATE TABLE IF NOT EXISTS journal_batches(
            id SERIAL PRIMARY KEY,
            batch_no VARCHAR(100) UNIQUE NOT NULL,
            batch_date DATE NOT NULL,
            description TEXT DEFAULT '',
            status VARCHAR(50) NOT NULL DEFAULT 'مسودة',
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS journal_batch_groups(
            id SERIAL PRIMARY KEY,
            batch_id INTEGER NOT NULL REFERENCES journal_batches(id) ON DELETE CASCADE,
            group_no INTEGER NOT NULL,
            journal_no VARCHAR(100),
            total_debit NUMERIC(18,2) NOT NULL DEFAULT 0,
            total_credit NUMERIC(18,2) NOT NULL DEFAULT 0,
            status VARCHAR(50) NOT NULL DEFAULT 'مسودة'
        )""",
        """CREATE TABLE IF NOT EXISTS journal_batch_lines(
            id SERIAL PRIMARY KEY,
            group_id INTEGER NOT NULL REFERENCES journal_batch_groups(id) ON DELETE CASCADE,
            account_id INTEGER NOT NULL REFERENCES chart_of_accounts(id),
            debit NUMERIC(18,2) NOT NULL DEFAULT 0,
            credit NUMERIC(18,2) NOT NULL DEFAULT 0,
            taxable INTEGER NOT NULL DEFAULT 0,
            tax_direction VARCHAR(50) DEFAULT 'غير مطبق',
            party_type VARCHAR(20) DEFAULT '',
            supplier_id INTEGER REFERENCES suppliers(id),
            customer_id INTEGER REFERENCES customers(id),
            tax_number VARCHAR(50) DEFAULT '',
            invoice_number VARCHAR(100) DEFAULT '',
            invoice_date DATE,
            line_description TEXT DEFAULT '',
            cost_center_id INTEGER REFERENCES cost_centers(id)
        )""",
        """CREATE TABLE IF NOT EXISTS fiscal_years(
            id SERIAL PRIMARY KEY,name VARCHAR(100) NOT NULL,start_date DATE NOT NULL,
            end_date DATE NOT NULL,status VARCHAR(30) NOT NULL DEFAULT 'مفتوحة',
            UNIQUE(start_date,end_date)
        )""",
        """CREATE TABLE IF NOT EXISTS fiscal_periods(
            id SERIAL PRIMARY KEY,fiscal_year_id INTEGER NOT NULL REFERENCES fiscal_years(id) ON DELETE CASCADE,
            name VARCHAR(100) NOT NULL,start_date DATE NOT NULL,end_date DATE NOT NULL,
            status VARCHAR(30) NOT NULL DEFAULT 'مفتوحة'
        )""",
        """CREATE TABLE IF NOT EXISTS treasury_vouchers(
            id SERIAL PRIMARY KEY,
            voucher_no VARCHAR(100) UNIQUE NOT NULL,
            voucher_type VARCHAR(30) NOT NULL,
            voucher_date DATE NOT NULL,
            party_type VARCHAR(20) DEFAULT '',
            customer_id INTEGER REFERENCES customers(id),
            supplier_id INTEGER REFERENCES suppliers(id),
            cash_bank_account_id INTEGER NOT NULL REFERENCES chart_of_accounts(id),
            counter_account_id INTEGER NOT NULL REFERENCES chart_of_accounts(id),
            amount NUMERIC(18,2) NOT NULL,
            payment_method VARCHAR(30) DEFAULT 'نقدي',
            reference VARCHAR(100) DEFAULT '',
            description TEXT DEFAULT '',
            cost_center_id INTEGER REFERENCES cost_centers(id),
            status VARCHAR(30) DEFAULT 'مسودة',
            posting_status VARCHAR(30) DEFAULT 'غير مرحّل',
            journal_id INTEGER,
            reversed_by INTEGER,
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS purchase_requisitions(
            id SERIAL PRIMARY KEY,
            requisition_no VARCHAR(100) UNIQUE NOT NULL,
            requisition_date DATE NOT NULL,
            branch_id INTEGER REFERENCES branches(id),
            cost_center_id INTEGER REFERENCES cost_centers(id),
            requested_by VARCHAR(255) DEFAULT '',
            department VARCHAR(255) DEFAULT '',
            priority VARCHAR(30) DEFAULT 'عادية',
            reason TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            status VARCHAR(50) DEFAULT 'مسودة',
            total_estimated NUMERIC(18,2) DEFAULT 0,
            approved_by INTEGER,
            approved_at TIMESTAMP,
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS purchase_requisition_items(
            id SERIAL PRIMARY KEY,
            requisition_id INTEGER NOT NULL REFERENCES purchase_requisitions(id) ON DELETE CASCADE,
            item_code VARCHAR(100) DEFAULT '',
            item_name VARCHAR(255) NOT NULL,
            description TEXT DEFAULT '',
            quantity NUMERIC(18,3) NOT NULL,
            unit VARCHAR(50) DEFAULT '',
            estimated_price NUMERIC(18,2) DEFAULT 0,
            required_date DATE,
            suggested_supplier_id INTEGER REFERENCES suppliers(id)
        )""",
        """CREATE TABLE IF NOT EXISTS purchase_orders(
            id SERIAL PRIMARY KEY,
            po_no VARCHAR(100) UNIQUE NOT NULL,
            po_date DATE NOT NULL,
            requisition_id INTEGER REFERENCES purchase_requisitions(id),
            supplier_id INTEGER NOT NULL REFERENCES suppliers(id),
            branch_id INTEGER REFERENCES branches(id),
            cost_center_id INTEGER REFERENCES cost_centers(id),
            payment_terms VARCHAR(255) DEFAULT '',
            delivery_terms VARCHAR(255) DEFAULT '',
            warehouse VARCHAR(255) DEFAULT '',
            notes TEXT DEFAULT '',
            status VARCHAR(50) DEFAULT 'مسودة',
            subtotal NUMERIC(18,2) DEFAULT 0,
            vat NUMERIC(18,2) DEFAULT 0,
            total NUMERIC(18,2) DEFAULT 0,
            approved_by INTEGER,
            approved_at TIMESTAMP,
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS purchase_order_items(
            id SERIAL PRIMARY KEY,
            po_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
            item_code VARCHAR(100) DEFAULT '',
            item_name VARCHAR(255) NOT NULL,
            description TEXT DEFAULT '',
            quantity NUMERIC(18,3) NOT NULL,
            received_qty NUMERIC(18,3) DEFAULT 0,
            unit VARCHAR(50) DEFAULT '',
            unit_price NUMERIC(18,2) NOT NULL,
            vat_rate NUMERIC(5,2) DEFAULT 15
        )""",
        """CREATE TABLE IF NOT EXISTS goods_receipts(
            id SERIAL PRIMARY KEY,
            grn_no VARCHAR(100) UNIQUE NOT NULL,
            grn_date DATE NOT NULL,
            po_id INTEGER NOT NULL REFERENCES purchase_orders(id),
            supplier_id INTEGER NOT NULL REFERENCES suppliers(id),
            warehouse VARCHAR(255) DEFAULT '',
            notes TEXT DEFAULT '',
            status VARCHAR(50) DEFAULT 'معتمد',
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS goods_receipt_items(
            id SERIAL PRIMARY KEY,
            grn_id INTEGER NOT NULL REFERENCES goods_receipts(id) ON DELETE CASCADE,
            po_item_id INTEGER NOT NULL REFERENCES purchase_order_items(id),
            received_qty NUMERIC(18,3) NOT NULL,
            accepted_qty NUMERIC(18,3) NOT NULL,
            rejected_qty NUMERIC(18,3) DEFAULT 0,
            notes TEXT DEFAULT ''
        )""",
        """CREATE TABLE IF NOT EXISTS supplier_invoices(
            id SERIAL PRIMARY KEY,
            supplier_invoice_no VARCHAR(100) NOT NULL,
            internal_no VARCHAR(100) UNIQUE NOT NULL,
            invoice_date DATE NOT NULL,
            supplier_id INTEGER NOT NULL REFERENCES suppliers(id),
            po_id INTEGER REFERENCES purchase_orders(id),
            grn_id INTEGER REFERENCES goods_receipts(id),
            cost_center_id INTEGER REFERENCES cost_centers(id),
            expense_or_inventory_account_id INTEGER NOT NULL REFERENCES chart_of_accounts(id),
            subtotal NUMERIC(18,2) NOT NULL,
            vat NUMERIC(18,2) DEFAULT 0,
            total NUMERIC(18,2) NOT NULL,
            status VARCHAR(50) DEFAULT 'مسودة',
            posting_status VARCHAR(30) DEFAULT 'غير مرحّل',
            journal_id INTEGER,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(supplier_id,supplier_invoice_no)
        )""",
        """CREATE TABLE IF NOT EXISTS approval_rules(
            id SERIAL PRIMARY KEY,
            document_type VARCHAR(50) NOT NULL,
            min_amount NUMERIC(18,2) DEFAULT 0,
            max_amount NUMERIC(18,2),
            approver_role VARCHAR(100) NOT NULL,
            active INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS warehouses(
            id SERIAL PRIMARY KEY,
            code VARCHAR(50) UNIQUE NOT NULL,
            name VARCHAR(255) NOT NULL,
            branch_id INTEGER REFERENCES branches(id),
            active INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS inventory_categories(
            id SERIAL PRIMARY KEY,
            code VARCHAR(50) UNIQUE,
            name VARCHAR(255) NOT NULL,
            parent_id INTEGER REFERENCES inventory_categories(id),
            active INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS inventory_movements(
            id SERIAL PRIMARY KEY,
            movement_no VARCHAR(100) UNIQUE NOT NULL,
            movement_date DATE NOT NULL,
            movement_type VARCHAR(50) NOT NULL,
            item_id INTEGER NOT NULL REFERENCES inventory(id),
            warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
            destination_warehouse_id INTEGER REFERENCES warehouses(id),
            quantity NUMERIC(18,3) NOT NULL,
            unit_cost NUMERIC(18,4) DEFAULT 0,
            total_cost NUMERIC(18,2) DEFAULT 0,
            reference_type VARCHAR(50) DEFAULT '',
            reference_id INTEGER,
            reference_no VARCHAR(100) DEFAULT '',
            batch_no VARCHAR(100) DEFAULT '',
            serial_no VARCHAR(100) DEFAULT '',
            production_date DATE,
            expiry_date DATE,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS inventory_counts(
            id SERIAL PRIMARY KEY,
            count_no VARCHAR(100) UNIQUE NOT NULL,
            count_date DATE NOT NULL,
            warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
            status VARCHAR(30) DEFAULT 'مسودة',
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS inventory_count_items(
            id SERIAL PRIMARY KEY,
            count_id INTEGER NOT NULL REFERENCES inventory_counts(id) ON DELETE CASCADE,
            item_id INTEGER NOT NULL REFERENCES inventory(id),
            system_qty NUMERIC(18,3) NOT NULL DEFAULT 0,
            counted_qty NUMERIC(18,3) NOT NULL DEFAULT 0,
            variance_qty NUMERIC(18,3) NOT NULL DEFAULT 0,
            unit_cost NUMERIC(18,4) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS sales_quotations(
            id SERIAL PRIMARY KEY, quotation_no VARCHAR(100) UNIQUE NOT NULL,
            quotation_date DATE NOT NULL, valid_until DATE,
            customer_id INTEGER NOT NULL REFERENCES customers(id),
            branch_id INTEGER REFERENCES branches(id), cost_center_id INTEGER REFERENCES cost_centers(id),
            status VARCHAR(40) DEFAULT 'مسودة', subtotal NUMERIC(18,2) DEFAULT 0,
            discount NUMERIC(18,2) DEFAULT 0, vat NUMERIC(18,2) DEFAULT 0,
            total NUMERIC(18,2) DEFAULT 0, notes TEXT DEFAULT '',
            converted_order_id INTEGER, created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS sales_quotation_items(
            id SERIAL PRIMARY KEY, quotation_id INTEGER NOT NULL REFERENCES sales_quotations(id) ON DELETE CASCADE,
            item_id INTEGER REFERENCES inventory(id), item_name VARCHAR(255) NOT NULL,
            quantity NUMERIC(18,3) NOT NULL, unit VARCHAR(50) DEFAULT 'وحدة',
            unit_price NUMERIC(18,2) NOT NULL, discount_rate NUMERIC(5,2) DEFAULT 0,
            vat_rate NUMERIC(5,2) DEFAULT 15, line_subtotal NUMERIC(18,2) DEFAULT 0,
            line_discount NUMERIC(18,2) DEFAULT 0, line_vat NUMERIC(18,2) DEFAULT 0,
            line_total NUMERIC(18,2) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS sales_orders(
            id SERIAL PRIMARY KEY, order_no VARCHAR(100) UNIQUE NOT NULL,
            order_date DATE NOT NULL, quotation_id INTEGER REFERENCES sales_quotations(id),
            customer_id INTEGER NOT NULL REFERENCES customers(id),
            branch_id INTEGER REFERENCES branches(id), cost_center_id INTEGER REFERENCES cost_centers(id),
            warehouse_id INTEGER REFERENCES warehouses(id), status VARCHAR(40) DEFAULT 'مسودة',
            subtotal NUMERIC(18,2) DEFAULT 0, discount NUMERIC(18,2) DEFAULT 0,
            vat NUMERIC(18,2) DEFAULT 0, total NUMERIC(18,2) DEFAULT 0,
            notes TEXT DEFAULT '', created_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS sales_order_items(
            id SERIAL PRIMARY KEY, order_id INTEGER NOT NULL REFERENCES sales_orders(id) ON DELETE CASCADE,
            item_id INTEGER REFERENCES inventory(id), item_name VARCHAR(255) NOT NULL,
            quantity NUMERIC(18,3) NOT NULL, delivered_qty NUMERIC(18,3) DEFAULT 0,
            unit VARCHAR(50) DEFAULT 'وحدة', unit_price NUMERIC(18,2) NOT NULL,
            discount_rate NUMERIC(5,2) DEFAULT 0, vat_rate NUMERIC(5,2) DEFAULT 15,
            line_subtotal NUMERIC(18,2) DEFAULT 0, line_discount NUMERIC(18,2) DEFAULT 0,
            line_vat NUMERIC(18,2) DEFAULT 0, line_total NUMERIC(18,2) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS sales_deliveries(
            id SERIAL PRIMARY KEY, delivery_no VARCHAR(100) UNIQUE NOT NULL,
            delivery_date DATE NOT NULL, order_id INTEGER NOT NULL REFERENCES sales_orders(id),
            customer_id INTEGER NOT NULL REFERENCES customers(id),
            warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
            status VARCHAR(40) DEFAULT 'معتمد', notes TEXT DEFAULT '',
            invoice_id INTEGER, created_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS sales_delivery_items(
            id SERIAL PRIMARY KEY, delivery_id INTEGER NOT NULL REFERENCES sales_deliveries(id) ON DELETE CASCADE,
            order_item_id INTEGER NOT NULL REFERENCES sales_order_items(id),
            item_id INTEGER REFERENCES inventory(id), quantity NUMERIC(18,3) NOT NULL,
            unit_cost NUMERIC(18,4) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS sales_returns(
            id SERIAL PRIMARY KEY,
            return_no VARCHAR(100) UNIQUE NOT NULL,
            return_date DATE NOT NULL,
            invoice_id INTEGER NOT NULL REFERENCES invoices(id),
            customer_id INTEGER NOT NULL REFERENCES customers(id),
            warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
            subtotal NUMERIC(18,2) DEFAULT 0,
            vat NUMERIC(18,2) DEFAULT 0,
            total NUMERIC(18,2) DEFAULT 0,
            cost_total NUMERIC(18,2) DEFAULT 0,
            status VARCHAR(40) DEFAULT 'معتمد',
            journal_id INTEGER,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS sales_return_items(
            id SERIAL PRIMARY KEY,
            return_id INTEGER NOT NULL REFERENCES sales_returns(id) ON DELETE CASCADE,
            invoice_item_id INTEGER REFERENCES invoice_items(id),
            item_id INTEGER REFERENCES inventory(id),
            item_name VARCHAR(255) NOT NULL,
            quantity NUMERIC(18,3) NOT NULL,
            unit VARCHAR(50) DEFAULT 'وحدة',
            unit_price NUMERIC(18,2) NOT NULL,
            vat_rate NUMERIC(5,2) DEFAULT 15,
            unit_cost NUMERIC(18,4) DEFAULT 0,
            line_subtotal NUMERIC(18,2) DEFAULT 0,
            line_vat NUMERIC(18,2) DEFAULT 0,
            line_total NUMERIC(18,2) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS invoice_payment_allocations(
            id SERIAL PRIMARY KEY,
            invoice_id INTEGER NOT NULL REFERENCES invoices(id),
            voucher_id INTEGER NOT NULL REFERENCES treasury_vouchers(id),
            allocated_amount NUMERIC(18,2) NOT NULL,
            allocation_date DATE NOT NULL,
            created_by INTEGER,
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(invoice_id,voucher_id)
        )""",
        """CREATE TABLE IF NOT EXISTS asset_categories(
            id SERIAL PRIMARY KEY,
            code VARCHAR(50) UNIQUE NOT NULL,
            name VARCHAR(255) NOT NULL,
            useful_life_months INTEGER NOT NULL DEFAULT 60,
            depreciation_method VARCHAR(50) DEFAULT 'القسط الثابت',
            asset_account_id INTEGER REFERENCES chart_of_accounts(id),
            accumulated_depr_account_id INTEGER REFERENCES chart_of_accounts(id),
            depreciation_expense_account_id INTEGER REFERENCES chart_of_accounts(id),
            active INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS fixed_assets(
            id SERIAL PRIMARY KEY,
            asset_no VARCHAR(100) UNIQUE NOT NULL,
            name VARCHAR(255) NOT NULL,
            name_en VARCHAR(255) DEFAULT '',
            category_id INTEGER REFERENCES asset_categories(id),
            branch_id INTEGER REFERENCES branches(id),
            cost_center_id INTEGER REFERENCES cost_centers(id),
            supplier_id INTEGER REFERENCES suppliers(id),
            purchase_date DATE NOT NULL,
            capitalization_date DATE NOT NULL,
            purchase_cost NUMERIC(18,2) NOT NULL DEFAULT 0,
            residual_value NUMERIC(18,2) DEFAULT 0,
            useful_life_months INTEGER NOT NULL DEFAULT 60,
            depreciation_method VARCHAR(50) DEFAULT 'القسط الثابت',
            accumulated_depreciation NUMERIC(18,2) DEFAULT 0,
            net_book_value NUMERIC(18,2) DEFAULT 0,
            serial_number VARCHAR(100) DEFAULT '',
            location VARCHAR(255) DEFAULT '',
            status VARCHAR(50) DEFAULT 'نشط',
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS asset_depreciation_runs(
            id SERIAL PRIMARY KEY,
            run_no VARCHAR(100) UNIQUE NOT NULL,
            period_date DATE NOT NULL,
            status VARCHAR(40) DEFAULT 'مسودة',
            total_amount NUMERIC(18,2) DEFAULT 0,
            journal_id INTEGER,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS asset_depreciation_lines(
            id SERIAL PRIMARY KEY,
            run_id INTEGER NOT NULL REFERENCES asset_depreciation_runs(id) ON DELETE CASCADE,
            asset_id INTEGER NOT NULL REFERENCES fixed_assets(id),
            depreciation_amount NUMERIC(18,2) NOT NULL DEFAULT 0,
            accumulated_before NUMERIC(18,2) DEFAULT 0,
            accumulated_after NUMERIC(18,2) DEFAULT 0,
            net_book_value_after NUMERIC(18,2) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS asset_disposals(
            id SERIAL PRIMARY KEY,
            disposal_no VARCHAR(100) UNIQUE NOT NULL,
            asset_id INTEGER NOT NULL REFERENCES fixed_assets(id),
            disposal_date DATE NOT NULL,
            disposal_type VARCHAR(50) NOT NULL,
            proceeds NUMERIC(18,2) DEFAULT 0,
            gain_loss NUMERIC(18,2) DEFAULT 0,
            journal_id INTEGER,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS departments(
            id SERIAL PRIMARY KEY,
            code VARCHAR(50) UNIQUE NOT NULL,
            name VARCHAR(255) NOT NULL,
            manager_employee_id INTEGER,
            active INTEGER DEFAULT 1
        )""",
        """CREATE TABLE IF NOT EXISTS payroll_settings(
            id INTEGER PRIMARY KEY,
            salary_expense_account_id INTEGER REFERENCES chart_of_accounts(id),
            payroll_payable_account_id INTEGER REFERENCES chart_of_accounts(id),
            deduction_account_id INTEGER REFERENCES chart_of_accounts(id),
            working_days_per_month NUMERIC(8,2) DEFAULT 30,
            working_hours_per_day NUMERIC(8,2) DEFAULT 8,
            overtime_multiplier NUMERIC(8,2) DEFAULT 1.5
        )""",
        """CREATE TABLE IF NOT EXISTS attendance_records(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id),
            attendance_date DATE NOT NULL,
            status VARCHAR(30) DEFAULT 'حاضر',
            check_in TIME,
            check_out TIME,
            overtime_hours NUMERIC(8,2) DEFAULT 0,
            absence_hours NUMERIC(8,2) DEFAULT 0,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(employee_id,attendance_date)
        )""",
        """CREATE TABLE IF NOT EXISTS payroll_runs(
            id SERIAL PRIMARY KEY,
            run_no VARCHAR(100) UNIQUE NOT NULL,
            period_start DATE NOT NULL,
            period_end DATE NOT NULL,
            payment_date DATE,
            status VARCHAR(40) DEFAULT 'مسودة',
            total_gross NUMERIC(18,2) DEFAULT 0,
            total_deductions NUMERIC(18,2) DEFAULT 0,
            total_net NUMERIC(18,2) DEFAULT 0,
            journal_id INTEGER,
            notes TEXT DEFAULT '',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS payroll_lines(
            id SERIAL PRIMARY KEY,
            run_id INTEGER NOT NULL REFERENCES payroll_runs(id) ON DELETE CASCADE,
            employee_id INTEGER NOT NULL REFERENCES employees(id),
            basic_salary NUMERIC(18,2) DEFAULT 0,
            allowances NUMERIC(18,2) DEFAULT 0,
            overtime_hours NUMERIC(8,2) DEFAULT 0,
            overtime_amount NUMERIC(18,2) DEFAULT 0,
            absence_days NUMERIC(8,2) DEFAULT 0,
            absence_deduction NUMERIC(18,2) DEFAULT 0,
            other_deductions NUMERIC(18,2) DEFAULT 0,
            gross_salary NUMERIC(18,2) DEFAULT 0,
            total_deductions NUMERIC(18,2) DEFAULT 0,
            net_salary NUMERIC(18,2) DEFAULT 0,
            cost_center_id INTEGER REFERENCES cost_centers(id),
            notes TEXT DEFAULT '',
            UNIQUE(run_id,employee_id)
        )""",
        """CREATE TABLE IF NOT EXISTS employee_salary_adjustments(
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL REFERENCES employees(id),
            adjustment_date DATE NOT NULL,
            adjustment_type VARCHAR(30) NOT NULL,
            amount NUMERIC(18,2) NOT NULL,
            recurring INTEGER DEFAULT 0,
            description TEXT DEFAULT '',
            active INTEGER DEFAULT 1,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS audit_logs(
            id SERIAL PRIMARY KEY,
            user_id INTEGER,
            username VARCHAR(100),
            action VARCHAR(100) NOT NULL,
            entity VARCHAR(100) NOT NULL,
            details TEXT DEFAULT '',
            created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""",
        """CREATE TABLE IF NOT EXISTS expenses(
            id SERIAL PRIMARY KEY,
            expense_date DATE NOT NULL,
            category VARCHAR(255) NOT NULL,
            description TEXT,
            amount NUMERIC(18,2) NOT NULL,
            vat NUMERIC(18,2) DEFAULT 0,
            total NUMERIC(18,2) NOT NULL,
            branch_id INTEGER REFERENCES branches(id)
        )""",
        """CREATE TABLE IF NOT EXISTS inventory(
            id SERIAL PRIMARY KEY,
            sku VARCHAR(100) UNIQUE,
            name VARCHAR(255) NOT NULL,
            quantity NUMERIC(18,3) DEFAULT 0,
            unit VARCHAR(50) DEFAULT 'وحدة',
            cost NUMERIC(18,2) DEFAULT 0,
            sale_price NUMERIC(18,2) DEFAULT 0,
            reorder_level NUMERIC(18,3) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS employees(
            id SERIAL PRIMARY KEY,
            employee_no VARCHAR(100) UNIQUE,
            name VARCHAR(255) NOT NULL,
            job_title VARCHAR(255),
            branch_id INTEGER REFERENCES branches(id),
            basic_salary NUMERIC(18,2) DEFAULT 0,
            allowances NUMERIC(18,2) DEFAULT 0,
            active INTEGER DEFAULT 1
        )"""
    ]
    for statement in statements:
        db.session.execute(text(statement))

    # Safe schema upgrades for existing Neon databases.
    migrations = [
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS address TEXT DEFAULT ''",
        "ALTER TABLE customers ADD COLUMN IF NOT EXISTS name_en VARCHAR(255) DEFAULT ''",
        "ALTER TABLE suppliers ADD COLUMN IF NOT EXISTS name_en VARCHAR(255) DEFAULT ''",
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS phone VARCHAR(50) DEFAULT ''",
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS email VARCHAR(255) DEFAULT ''",
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS logo_url TEXT DEFAULT ''",
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS logo_data TEXT DEFAULT ''",
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS logo_mime VARCHAR(50) DEFAULT ''",
        "ALTER TABLE settings ADD COLUMN IF NOT EXISTS vat_rate NUMERIC(5,2) DEFAULT 15.00",
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS invoice_uuid VARCHAR(64)",
        "ALTER TABLE journal_entry_lines ADD COLUMN IF NOT EXISTS invoice_date DATE",
        "ALTER TABLE journal_entry_lines ADD COLUMN IF NOT EXISTS customer_id INTEGER REFERENCES customers(id)",
        "ALTER TABLE journal_entry_lines ADD COLUMN IF NOT EXISTS party_type VARCHAR(20) DEFAULT ''",
        "ALTER TABLE journal_entry_lines ADD COLUMN IF NOT EXISTS tax_number VARCHAR(50) DEFAULT ''",
        "ALTER TABLE chart_of_accounts ADD COLUMN IF NOT EXISTS normal_balance VARCHAR(20) DEFAULT 'مدين'",
        "ALTER TABLE chart_of_accounts ADD COLUMN IF NOT EXISTS statement_type VARCHAR(50) DEFAULT 'الميزانية العمومية'",
        'ALTER TABLE invoice_items ADD COLUMN IF NOT EXISTS item_id INTEGER',
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS sales_order_id INTEGER',
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS delivery_id INTEGER',
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS cost_center_id INTEGER',
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS warehouse_id INTEGER',
        "ALTER TABLE employees ADD COLUMN IF NOT EXISTS name_en VARCHAR(255) DEFAULT ''",
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS department_id INTEGER',
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS cost_center_id INTEGER',
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS hire_date DATE',
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS contract_end_date DATE',
        "ALTER TABLE employees ADD COLUMN IF NOT EXISTS iqama_no VARCHAR(100) DEFAULT ''",
        "ALTER TABLE employees ADD COLUMN IF NOT EXISTS bank_iban VARCHAR(100) DEFAULT ''",
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS housing_allowance NUMERIC(18,2) DEFAULT 0',
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS transport_allowance NUMERIC(18,2) DEFAULT 0',
        'ALTER TABLE employees ADD COLUMN IF NOT EXISTS other_allowance NUMERIC(18,2) DEFAULT 0',
        'ALTER TABLE sales_returns ADD COLUMN IF NOT EXISTS cost_total NUMERIC(18,2) DEFAULT 0',
        'ALTER TABLE sales_returns ADD COLUMN IF NOT EXISTS journal_id INTEGER',
        "ALTER TABLE sales_returns ADD COLUMN IF NOT EXISTS status VARCHAR(40) DEFAULT 'معتمد'",
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS payment_status VARCHAR(30) DEFAULT \'غير مسددة\'',
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS payment_terms VARCHAR(255) DEFAULT ''",
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS payment_method VARCHAR(50) DEFAULT ''",
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS sales_person VARCHAR(255) DEFAULT ''",
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS customer_reference VARCHAR(100) DEFAULT ''",
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS discount NUMERIC(18,2) DEFAULT 0",
        "ALTER TABLE sales_quotations ADD COLUMN IF NOT EXISTS sales_person VARCHAR(255) DEFAULT ''",
        "ALTER TABLE sales_quotations ADD COLUMN IF NOT EXISTS payment_terms VARCHAR(255) DEFAULT ''",
        "ALTER TABLE sales_orders ADD COLUMN IF NOT EXISTS sales_person VARCHAR(255) DEFAULT ''",
        "ALTER TABLE sales_orders ADD COLUMN IF NOT EXISTS payment_terms VARCHAR(255) DEFAULT ''",
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS due_date DATE',
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS payment_status VARCHAR(30) DEFAULT 'غير مسددة'",
        "ALTER TABLE inventory ADD COLUMN IF NOT EXISTS name_en VARCHAR(255) DEFAULT ''",
        "ALTER TABLE inventory ADD COLUMN IF NOT EXISTS barcode VARCHAR(100) DEFAULT ''",
        "ALTER TABLE inventory ADD COLUMN IF NOT EXISTS qr_code VARCHAR(255) DEFAULT ''",
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS category_id INTEGER',
        "ALTER TABLE inventory ADD COLUMN IF NOT EXISTS item_type VARCHAR(30) DEFAULT 'مخزني'",
        "ALTER TABLE inventory ADD COLUMN IF NOT EXISTS valuation_method VARCHAR(30) DEFAULT 'متوسط متحرك'",
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS min_level NUMERIC(18,3) DEFAULT 0',
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS max_level NUMERIC(18,3) DEFAULT 0',
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS track_batch INTEGER DEFAULT 0',
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS track_serial INTEGER DEFAULT 0',
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS track_expiry INTEGER DEFAULT 0',
        'ALTER TABLE inventory ADD COLUMN IF NOT EXISTS active INTEGER DEFAULT 1',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS customer_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS supplier_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS sales_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS purchases_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS vat_input_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS vat_output_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS cash_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS bank_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS inventory_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS cost_of_sales_account_id INTEGER',
        'ALTER TABLE settings ADD COLUMN IF NOT EXISTS retained_earnings_account_id INTEGER',
        'ALTER TABLE invoices ADD COLUMN IF NOT EXISTS journal_id INTEGER',
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS posting_status VARCHAR(30) DEFAULT 'غير مرحّل'",
        "ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS source_type VARCHAR(50) DEFAULT 'MANUAL'",
        'ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS source_id INTEGER',
        'ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS posted_at TIMESTAMP',
        'ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS reversal_of INTEGER',
        'ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS reversed_by INTEGER',
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS supplier_id INTEGER',
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS expense_account_id INTEGER',
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS payment_account_id INTEGER',
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS cost_center_id INTEGER',
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS taxable INTEGER DEFAULT 0',
        "ALTER TABLE expenses ADD COLUMN IF NOT EXISTS invoice_number VARCHAR(100) DEFAULT ''",
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS invoice_date DATE',
        "ALTER TABLE expenses ADD COLUMN IF NOT EXISTS status VARCHAR(30) DEFAULT 'مسودة'",
        'ALTER TABLE expenses ADD COLUMN IF NOT EXISTS journal_id INTEGER',
        "ALTER TABLE expenses ADD COLUMN IF NOT EXISTS posting_status VARCHAR(30) DEFAULT 'غير مرحّل'"
    ]
    for migration in migrations:
        db.session.execute(text(migration))

    default_password_hash = generate_password_hash("admin123")
    db.session.execute(
        text("""
            INSERT INTO users(id, username, password, role)
            VALUES(1,'admin',:password,'admin')
            ON CONFLICT (id) DO NOTHING
        """),
        {"password": default_password_hash}
    )
    db.session.execute(text("""
        INSERT INTO settings(id)
        VALUES(1)
        ON CONFLICT (id) DO NOTHING
    """))
    db.session.execute(text("""
        INSERT INTO branches(id,name,city,active)
        VALUES(1,'الفرع الرئيسي','الدمام',1)
        ON CONFLICT (id) DO NOTHING
    """))

    defaults = [
        ("1000","الأصول","Assets","أصل",None,1,0),
        ("1100","الأصول المتداولة","Current Assets","أصل","1000",2,0),
        ("1110","الصندوق","Cash","أصل","1100",3,1),
        ("1120","البنوك","Banks","أصل","1100",3,1),
        ("1130","العملاء","Accounts Receivable","أصل","1100",3,1),
        ("1140","ضريبة القيمة المضافة - مدخلات","VAT Input","أصل","1100",3,1),
        ("1200","المخزون","Inventory","أصل","1000",2,1),
        ("2000","الخصوم","Liabilities","خصم",None,1,0),
        ("2100","الموردون","Accounts Payable","خصم","2000",2,1),
        ("2200","ضريبة القيمة المضافة - مخرجات","VAT Output","خصم","2000",2,1),
        ("3000","حقوق الملكية","Equity","حقوق ملكية",None,1,0),
        ("3100","رأس المال","Capital","حقوق ملكية","3000",2,1),
        ("4000","الإيرادات","Revenue","إيراد",None,1,0),
        ("4100","إيرادات المبيعات","Sales Revenue","إيراد","4000",2,1),
        ("5000","تكلفة المبيعات والمشاريع","Cost of Sales","مصروف",None,1,0),
        ("5100","مواد المشاريع","Project Materials","مصروف","5000",2,1),
        ("6000","المصروفات","Expenses","مصروف",None,1,0),
        ("6100","مصروفات إدارية","Administrative Expenses","مصروف","6000",2,1),
    ]
    ids = {}
    for code, ar, en, typ, parent_code, level, accepts in defaults:
        parent_id = ids.get(parent_code)
        existing = db.session.execute(text("SELECT id FROM chart_of_accounts WHERE account_code=:c"), {"c": code}).scalar()
        if existing:
            ids[code] = existing
        else:
            new_id = db.session.execute(
                text("""INSERT INTO chart_of_accounts(
                    account_code,account_name_ar,account_name_en,account_type,
                    parent_id,level,accepts_entries,active)
                    VALUES(:c,:ar,:en,:t,:p,:l,:a,1) RETURNING id"""),
                {"c":code,"ar":ar,"en":en,"t":typ,"p":parent_id,"l":level,"a":accepts}
            ).scalar_one()
            ids[code] = new_id
    db.session.execute(text("""
        INSERT INTO cost_centers(code,name,active)
        VALUES('CC-001','الإدارة العامة',1)
        ON CONFLICT (code) DO NOTHING
    """))
    db.session.execute(text("""
        INSERT INTO warehouses(code,name,active)
        VALUES('WH-001','المستودع الرئيسي',1)
        ON CONFLICT (code) DO NOTHING
    """))

    for min_amount,max_amount,role in [
        (0,5000,"مدير القسم"),
        (5000.01,20000,"المدير المالي"),
        (20000.01,None,"المدير العام")
    ]:
        db.session.execute(text("""INSERT INTO approval_rules(
            document_type,min_amount,max_amount,approver_role,active)
            SELECT 'طلب شراء',:mn,:mx,:role,1
            WHERE NOT EXISTS(
              SELECT 1 FROM approval_rules WHERE document_type='طلب شراء'
              AND min_amount=:mn AND max_amount IS NOT DISTINCT FROM :mx
            )"""), {"mn":min_amount,"mx":max_amount,"role":role})

    current_year = datetime.now().year
    fy_id = db.session.execute(text("""INSERT INTO fiscal_years(name,start_date,end_date,status)
        VALUES(:n,:s,:e,'مفتوحة') ON CONFLICT(start_date,end_date)
        DO UPDATE SET name=EXCLUDED.name RETURNING id"""),
        {"n":f"السنة المالية {current_year}","s":f"{current_year}-01-01","e":f"{current_year}-12-31"}).scalar_one()
    import calendar
    for m in range(1,13):
        last=calendar.monthrange(current_year,m)[1]
        db.session.execute(text("""INSERT INTO fiscal_periods(fiscal_year_id,name,start_date,end_date,status)
            SELECT :fy,:n,:s,:e,'مفتوحة' WHERE NOT EXISTS(
            SELECT 1 FROM fiscal_periods WHERE start_date=:s AND end_date=:e)"""),
            {"fy":fy_id,"n":f"{current_year}-{m:02d}","s":f"{current_year}-{m:02d}-01",
             "e":f"{current_year}-{m:02d}-{last:02d}"})
    db.session.execute(text("""UPDATE settings SET
      customer_account_id=COALESCE(customer_account_id,:c),
      supplier_account_id=COALESCE(supplier_account_id,:s),
      sales_account_id=COALESCE(sales_account_id,:sales),
      purchases_account_id=COALESCE(purchases_account_id,:p),
      vat_input_account_id=COALESCE(vat_input_account_id,:vi),
      vat_output_account_id=COALESCE(vat_output_account_id,:vo),
      cash_account_id=COALESCE(cash_account_id,:cash),
      bank_account_id=COALESCE(bank_account_id,:bank),
      inventory_account_id=COALESCE(inventory_account_id,:inv),
      cost_of_sales_account_id=COALESCE(cost_of_sales_account_id,:cos),
      retained_earnings_account_id=COALESCE(retained_earnings_account_id,:re)
      WHERE id=1"""),{"c":ids.get("1130"),"s":ids.get("2100"),"sales":ids.get("4100"),
      "p":ids.get("5100"),"vi":ids.get("1140"),"vo":ids.get("2200"),"cash":ids.get("1110"),
      "bank":ids.get("1120"),"inv":ids.get("1200"),"cos":ids.get("5000"),"re":ids.get("3100")})

    db.session.commit()

@app.before_request
def ensure_database():
    if not getattr(app, "_db_initialized", False):
        init_db()
        app._db_initialized = True

@app.context_processor
def inject_settings():
    settings = row("SELECT * FROM settings WHERE id=1")
    return {"app_settings": settings, "app_version": APP_VERSION}

@app.route("/health")
def health():
    db.session.execute(text("SELECT 1"))
    return {"status": "ok"}, 200

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()
        user = row("SELECT * FROM users WHERE username=:u", {"u": username})

        password_ok = False
        if user:
            stored_password = user["password"] or ""
            if stored_password.startswith(("pbkdf2:", "scrypt:")):
                password_ok = check_password_hash(stored_password, password)
            else:
                # Automatic migration for the old plaintext admin password.
                password_ok = stored_password == password
                if password_ok:
                    execute(
                        "UPDATE users SET password=:password WHERE id=:id",
                        {"password": generate_password_hash(password), "id": user["id"]},
                    )

        if user and password_ok:
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            audit("LOGIN", "USER", f"تسجيل دخول المستخدم {username}")
            return redirect(url_for("dashboard"))

        flash("بيانات الدخول غير صحيحة", "danger")

    return render_template("login.html")

@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        current_password = request.form["current_password"]
        new_password = request.form["new_password"]
        confirm_password = request.form["confirm_password"]

        if len(new_password) < 8:
            flash("كلمة المرور الجديدة يجب أن تكون 8 أحرف على الأقل", "danger")
            return redirect(url_for("change_password"))

        if new_password != confirm_password:
            flash("تأكيد كلمة المرور غير مطابق", "danger")
            return redirect(url_for("change_password"))

        user = row("SELECT * FROM users WHERE id=:id", {"id": session["user_id"]})
        stored_password = user["password"] or ""
        valid_current = (
            check_password_hash(stored_password, current_password)
            if stored_password.startswith(("pbkdf2:", "scrypt:"))
            else stored_password == current_password
        )

        if not valid_current:
            flash("كلمة المرور الحالية غير صحيحة", "danger")
            return redirect(url_for("change_password"))

        execute(
            "UPDATE users SET password=:password WHERE id=:id",
            {"password": generate_password_hash(new_password), "id": session["user_id"]},
        )
        audit("UPDATE", "USER", "تم تغيير كلمة المرور")
        flash("تم تغيير كلمة المرور بنجاح", "success")
        return redirect(url_for("dashboard"))

    return render_template("change_password.html")


@app.route("/audit-logs")
@login_required
def audit_logs():
    logs = rows("""
        SELECT * FROM audit_logs
        ORDER BY id DESC
        LIMIT 300
    """)
    return render_template("audit_logs.html", logs=logs)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    sales = row("SELECT COALESCE(SUM(total),0) s FROM invoices WHERE status='معتمدة'")["s"]
    expenses_total = row("SELECT COALESCE(SUM(total),0) s FROM expenses")["s"]
    customers_count = row("SELECT COUNT(*) c FROM customers")["c"]
    employees_count = row("SELECT COUNT(*) c FROM employees WHERE active=1")["c"]
    recent = rows("""
        SELECT i.*, c.name customer_name FROM invoices i
        JOIN customers c ON c.id=i.customer_id
        ORDER BY i.id DESC LIMIT 5
    """)
    return render_template("dashboard.html", sales=float(sales), expenses=float(expenses_total),
                           customers=customers_count, employees=employees_count, recent=recent)

@app.route("/settings", methods=["GET","POST"])
@login_required
def settings():
    current = row("SELECT * FROM settings WHERE id=1")

    if request.method == "POST":
        logo_data = current["logo_data"] or ""
        logo_mime = current["logo_mime"] or ""

        uploaded_logo = request.files.get("logo_file")
        remove_logo = request.form.get("remove_logo") == "1"

        if remove_logo:
            logo_data = ""
            logo_mime = ""
        elif uploaded_logo and uploaded_logo.filename:
            allowed_mimes = {"image/png", "image/jpeg", "image/webp"}
            if uploaded_logo.mimetype not in allowed_mimes:
                flash("صيغة الشعار غير مدعومة. استخدم PNG أو JPG أو WEBP.", "danger")
                return redirect(url_for("settings"))

            raw = uploaded_logo.read()
            if len(raw) > 2 * 1024 * 1024:
                flash("حجم الشعار يجب ألا يتجاوز 2 ميجابايت.", "danger")
                return redirect(url_for("settings"))

            logo_data = base64.b64encode(raw).decode("ascii")
            logo_mime = uploaded_logo.mimetype

        execute("""UPDATE settings
                   SET company_name_ar=:ar, company_name_en=:en, vat_number=:vat,
                       cr_number=:cr, currency=:cur, address=:address, phone=:phone,
                       email=:email, logo_url=:logo_url, logo_data=:logo_data,
                       logo_mime=:logo_mime, vat_rate=:vat_rate,
                       customer_account_id=:customer_account_id,supplier_account_id=:supplier_account_id,
                       sales_account_id=:sales_account_id,purchases_account_id=:purchases_account_id,
                       vat_input_account_id=:vat_input_account_id,vat_output_account_id=:vat_output_account_id,
                       cash_account_id=:cash_account_id,bank_account_id=:bank_account_id,
                       inventory_account_id=:inventory_account_id,cost_of_sales_account_id=:cost_of_sales_account_id,
                       retained_earnings_account_id=:retained_earnings_account_id
                   WHERE id=1""",
                {"ar": request.form["company_name_ar"],
                 "en": request.form["company_name_en"],
                 "vat": request.form["vat_number"],
                 "cr": request.form["cr_number"],
                 "cur": request.form["currency"],
                 "address": request.form.get("address", ""),
                 "phone": request.form.get("phone", ""),
                 "email": request.form.get("email", ""),
                 "logo_url": request.form.get("logo_url", ""),
                 "logo_data": logo_data,
                 "logo_mime": logo_mime,
                 "vat_rate": float(request.form.get("vat_rate", 15) or 15),
                 "customer_account_id":request.form.get("customer_account_id") or None,
                 "supplier_account_id":request.form.get("supplier_account_id") or None,
                 "sales_account_id":request.form.get("sales_account_id") or None,
                 "purchases_account_id":request.form.get("purchases_account_id") or None,
                 "vat_input_account_id":request.form.get("vat_input_account_id") or None,
                 "vat_output_account_id":request.form.get("vat_output_account_id") or None,
                 "cash_account_id":request.form.get("cash_account_id") or None,
                 "bank_account_id":request.form.get("bank_account_id") or None,
                 "inventory_account_id":request.form.get("inventory_account_id") or None,
                 "cost_of_sales_account_id":request.form.get("cost_of_sales_account_id") or None,
                 "retained_earnings_account_id":request.form.get("retained_earnings_account_id") or None})

        audit("UPDATE", "SETTINGS", "تم تحديث إعدادات الشركة والشعار")
        flash("تم حفظ إعدادات الشركة والشعار", "success")
        return redirect(url_for("settings"))

    return render_template("settings.html",row=current,
        accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                         WHERE active=1 AND accepts_entries=1 ORDER BY account_code"""))

@app.route("/branches", methods=["GET","POST"])
@login_required
def branches():
    if request.method == "POST":
        execute("INSERT INTO branches(name,city) VALUES(:n,:c)",
                {"n": request.form["name"], "c": request.form["city"]})
        audit("CREATE", "BRANCH", f"إضافة فرع: {request.form['name']}")
        flash("تمت إضافة الفرع", "success")
    return render_template("branches.html", rows=rows("SELECT * FROM branches ORDER BY id DESC"))

@app.route("/customers", methods=["GET","POST"])
@login_required
def customers():
    if request.method == "POST":
        name_ar = request.form["name"].strip()
        name_en = request.form.get("name_en","").strip() or transliterate_arabic_name(name_ar)
        execute("""INSERT INTO customers(name,name_en,vat_number,phone,email)
                   VALUES(:n,:ne,:v,:p,:e)""",
                {"n":name_ar,"ne":name_en,"v":request.form["vat_number"],
                 "p":request.form["phone"],"e":request.form["email"]})
        audit("CREATE","CUSTOMER",f"إضافة عميل: {name_ar} / {name_en}")
        flash("تمت إضافة العميل", "success")
    return render_template("customers.html", rows=rows("SELECT * FROM customers ORDER BY id DESC"))

@app.route("/suppliers", methods=["GET","POST"])
@login_required
def suppliers():
    if request.method == "POST":
        name_ar = request.form["name"].strip()
        name_en = request.form.get("name_en","").strip() or transliterate_arabic_name(name_ar)
        execute("""INSERT INTO suppliers(name,name_en,vat_number,phone,email)
                   VALUES(:n,:ne,:v,:p,:e)""",
                {"n":name_ar,"ne":name_en,"v":request.form["vat_number"],
                 "p":request.form["phone"],"e":request.form["email"]})
        audit("CREATE","SUPPLIER",f"إضافة مورد: {name_ar} / {name_en}")
        flash("تمت إضافة المورد", "success")
    return render_template("suppliers.html", rows=rows("SELECT * FROM suppliers ORDER BY id DESC"))

@app.route("/invoices", methods=["GET","POST"])
@login_required
def invoices():
    if request.method == "POST":
        item_names = request.form.getlist("item_name[]")
        descriptions = request.form.getlist("description[]")
        quantities = request.form.getlist("quantity[]")
        units = request.form.getlist("unit[]")
        unit_prices = request.form.getlist("unit_price[]")

        company_settings = row("SELECT vat_rate FROM settings WHERE id=1")
        vat_rate = float(company_settings["vat_rate"] or 15)

        prepared_items = []
        invoice_subtotal = invoice_vat = invoice_total = 0.0

        for index, item_name in enumerate(item_names):
            item_name = (item_name or "").strip()
            if not item_name:
                continue

            description = descriptions[index].strip() if index < len(descriptions) else ""
            quantity = float(quantities[index] or 0)
            unit = units[index].strip() if index < len(units) else "وحدة"
            unit_price = float(unit_prices[index] or 0)

            if quantity <= 0:
                flash(f"الكمية في البند رقم {index + 1} يجب أن تكون أكبر من صفر", "danger")
                return redirect(url_for("invoices"))

            line_subtotal = round(quantity * unit_price, 2)
            line_vat = round(line_subtotal * vat_rate / 100, 2)
            line_total = round(line_subtotal + line_vat, 2)

            prepared_items.append({
                "item_name": item_name,
                "description": description,
                "quantity": quantity,
                "unit": unit or "وحدة",
                "unit_price": unit_price,
                "vat_rate": vat_rate,
                "line_subtotal": line_subtotal,
                "line_vat": line_vat,
                "line_total": line_total,
            })

            invoice_subtotal += line_subtotal
            invoice_vat += line_vat
            invoice_total += line_total

        if not prepared_items:
            flash("يجب إضافة بند صحيح واحد على الأقل", "danger")
            return redirect(url_for("invoices"))

        invoice_no = next_invoice_number(request.form["invoice_date"])
        invoice_uuid = str(uuid.uuid4())

        execute("""INSERT INTO invoices(
                      invoice_no,invoice_uuid,customer_id,invoice_date,
                      subtotal,vat,total,status,branch_id,notes,created_at
                   )
                   VALUES(
                      :no,:uuid,:cid,:dt,:sub,:vat,:tot,:st,:bid,:notes,:created
                   )""",
                {"no": invoice_no, "uuid": invoice_uuid,
                 "cid": request.form["customer_id"],
                 "dt": request.form["invoice_date"],
                 "sub": round(invoice_subtotal, 2),
                 "vat": round(invoice_vat, 2),
                 "tot": round(invoice_total, 2),
                 "st": request.form["status"],
                 "bid": request.form.get("branch_id") or None,
                 "notes": request.form.get("notes", ""),
                 "created": datetime.now()})

        invoice_id = row("SELECT id FROM invoices WHERE invoice_no=:no", {"no": invoice_no})["id"]

        for item in prepared_items:
            execute("""INSERT INTO invoice_items(
                          invoice_id,item_name,description,quantity,unit,unit_price,
                          vat_rate,line_subtotal,line_vat,line_total
                       )
                       VALUES(
                          :invoice_id,:item_name,:description,:quantity,:unit,:unit_price,
                          :vat_rate,:line_subtotal,:line_vat,:line_total
                       )""", {"invoice_id": invoice_id, **item})

        audit("CREATE","INVOICE",f"إنشاء فاتورة: {invoice_no}")
        if request.form["status"]=="معتمدة":
            try:
                post_invoice_to_ledger(invoice_id)
                flash("تم إنشاء الفاتورة وترحيلها محاسبيًا","success")
            except Exception as exc:
                db.session.rollback()
                execute("UPDATE invoices SET status='مسودة',posting_status='خطأ' WHERE id=:id",{"id":invoice_id})
                flash(f"تم حفظ الفاتورة كمسودة ولم تُرحّل: {exc}","danger")
        else:
            flash("تم إنشاء الفاتورة كمسودة","success")
        if request.form.get("print_after_save") == "1":
            return redirect(url_for("invoice_view", invoice_id=invoice_id, print=1))
        return redirect(url_for("invoice_view", invoice_id=invoice_id))

    invoice_rows = rows("""SELECT i.*, c.name customer_name, b.name branch_name
                           FROM invoices i
                           JOIN customers c ON c.id=i.customer_id
                           LEFT JOIN branches b ON b.id=i.branch_id
                           ORDER BY i.id DESC""")

    return render_template(
        "invoices.html",
        rows=invoice_rows,
        customers=rows("SELECT * FROM customers ORDER BY name"),
        branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"),
        inventory_items=rows("SELECT * FROM inventory ORDER BY name")
    )

@app.route("/expenses", methods=["GET","POST"])
@login_required
def expenses():
    if request.method=="POST":
        amount=round(float(request.form["amount"] or 0),2)
        taxable=1 if request.form.get("taxable")=="1" else 0
        rate=float(row("SELECT vat_rate FROM settings WHERE id=1")["vat_rate"] or 15)
        vat=round(amount*rate/100,2) if taxable else 0
        total=round(amount+vat,2)
        execute("""INSERT INTO expenses(expense_date,category,description,amount,vat,total,
          branch_id,supplier_id,expense_account_id,payment_account_id,cost_center_id,taxable,
          invoice_number,invoice_date,status,posting_status)
          VALUES(:dt,:cat,:des,:amt,:vat,:tot,:bid,:sid,:ea,:pa,:cc,:tax,:inv,:idate,:status,'غير مرحّل')""",
          {"dt":request.form["expense_date"],"cat":request.form["category"],
           "des":request.form.get("description",""),"amt":amount,"vat":vat,"tot":total,
           "bid":request.form.get("branch_id") or None,"sid":request.form.get("supplier_id") or None,
           "ea":request.form.get("expense_account_id") or None,"pa":request.form.get("payment_account_id") or None,
           "cc":request.form.get("cost_center_id") or None,"tax":taxable,
           "inv":request.form.get("invoice_number",""),"idate":request.form.get("invoice_date") or None,
           "status":request.form.get("status","مسودة")})
        expense_id=row("SELECT MAX(id) id FROM expenses")["id"]
        if request.form.get("status")=="معتمدة":
            try:
                post_expense_to_ledger(expense_id)
                flash("تم تسجيل المصروف وترحيله محاسبيًا","success")
            except Exception as exc:
                db.session.rollback()
                execute("UPDATE expenses SET status='مسودة',posting_status='خطأ' WHERE id=:id",{"id":expense_id})
                flash(f"تم حفظ المصروف كمسودة ولم يُرحّل: {exc}","danger")
        else: flash("تم حفظ المصروف كمسودة","success")
        audit("CREATE","EXPENSE",f"إضافة مصروف: {request.form['category']}")
        return redirect(url_for("expenses"))
    expense_rows=rows("""SELECT e.*,b.name branch_name,s.name supplier_name,
      a.account_name_ar expense_account_name,p.account_name_ar payment_account_name,
      cc.name cost_center_name,j.journal_no FROM expenses e
      LEFT JOIN branches b ON b.id=e.branch_id LEFT JOIN suppliers s ON s.id=e.supplier_id
      LEFT JOIN chart_of_accounts a ON a.id=e.expense_account_id
      LEFT JOIN chart_of_accounts p ON p.id=e.payment_account_id
      LEFT JOIN cost_centers cc ON cc.id=e.cost_center_id
      LEFT JOIN journal_entries j ON j.id=e.journal_id ORDER BY e.id DESC""")
    accts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                  WHERE active=1 AND accepts_entries=1 ORDER BY account_code""")
    return render_template("expenses.html",rows=expense_rows,
      branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"),
      suppliers=rows("SELECT * FROM suppliers ORDER BY name"),accounts=accts,
      centers=rows("SELECT * FROM cost_centers WHERE active=1 ORDER BY code"))


@app.route("/inventory", methods=["GET","POST"])
@login_required
def inventory():
    if request.method=="POST":
        sku=request.form.get("sku","").strip() or None
        name=request.form["name"].strip()
        execute("""INSERT INTO inventory(
          sku,name,name_en,barcode,category_id,item_type,unit,cost,sale_price,reorder_level,
          min_level,max_level,valuation_method,track_batch,track_serial,track_expiry,active,quantity)
          VALUES(:sku,:name,:name_en,:barcode,:category,:type,:unit,:cost,:sale,:reorder,
          :min,:max,:valuation,:batch,:serial,:expiry,1,0)""",
          {"sku":sku,"name":name,"name_en":request.form.get("name_en",""),
           "barcode":request.form.get("barcode",""),"category":request.form.get("category_id") or None,
           "type":request.form.get("item_type","مخزني"),"unit":request.form.get("unit","وحدة"),
           "cost":float(request.form.get("cost") or 0),"sale":float(request.form.get("sale_price") or 0),
           "reorder":float(request.form.get("reorder_level") or 0),
           "min":float(request.form.get("min_level") or 0),"max":float(request.form.get("max_level") or 0),
           "valuation":request.form.get("valuation_method","متوسط متحرك"),
           "batch":1 if request.form.get("track_batch") else 0,
           "serial":1 if request.form.get("track_serial") else 0,
           "expiry":1 if request.form.get("track_expiry") else 0})
        item_id=row("SELECT id FROM inventory WHERE sku IS NOT DISTINCT FROM :sku AND name=:name ORDER BY id DESC LIMIT 1",
                    {"sku":sku,"name":name})["id"]
        opening=float(request.form.get("opening_quantity") or 0)
        if opening>0:
            warehouse_id=request.form.get("warehouse_id",type=int)
            if not warehouse_id:
                flash("اختر المستودع للرصيد الافتتاحي","danger")
                return redirect(url_for("inventory"))
            record_inventory_movement(
                request.form.get("opening_date") or datetime.now().date().isoformat(),
                "رصيد افتتاحي",item_id,warehouse_id,opening,
                float(request.form.get("cost") or 0),notes="رصيد افتتاحي"
            )
        audit("CREATE","INVENTORY",f"إضافة صنف: {name}")
        flash("تم إنشاء بطاقة الصنف","success")
        return redirect(url_for("inventory_item_view",item_id=item_id))
    item_rows=rows("""SELECT i.*,c.name category_name,
      COALESCE((SELECT SUM(CASE
        WHEN m.movement_type IN ('رصيد افتتاحي','استلام','تسوية زيادة','تحويل وارد','مرتجع مبيعات') THEN m.quantity
        WHEN m.movement_type IN ('صرف','تسوية نقص','تحويل صادر','بيع','مرتجع مشتريات') THEN -m.quantity ELSE 0 END)
        FROM inventory_movements m WHERE m.item_id=i.id),0) stock_qty
      FROM inventory i LEFT JOIN inventory_categories c ON c.id=i.category_id
      ORDER BY i.id DESC""")
    return render_template("inventory.html",rows=item_rows,
      warehouses=rows("SELECT * FROM warehouses WHERE active=1 ORDER BY code"),
      categories=rows("SELECT * FROM inventory_categories WHERE active=1 ORDER BY name"))

@app.route("/inventory/items/<int:item_id>")
@login_required
def inventory_item_view(item_id):
    item=row("""SELECT i.*,c.name category_name FROM inventory i
                LEFT JOIN inventory_categories c ON c.id=i.category_id WHERE i.id=:id""",{"id":item_id})
    if not item:return "الصنف غير موجود",404
    balances=rows("""SELECT w.id,w.code,w.name,
      COALESCE(SUM(CASE
        WHEN m.movement_type IN ('رصيد افتتاحي','استلام','تسوية زيادة','تحويل وارد','مرتجع مبيعات') THEN m.quantity
        WHEN m.movement_type IN ('صرف','تسوية نقص','تحويل صادر','بيع','مرتجع مشتريات') THEN -m.quantity ELSE 0 END),0) quantity
      FROM warehouses w LEFT JOIN inventory_movements m ON m.warehouse_id=w.id AND m.item_id=:item
      GROUP BY w.id,w.code,w.name ORDER BY w.code""",{"item":item_id})
    movements=rows("""SELECT m.*,w.code warehouse_code,w.name warehouse_name,
                      d.code destination_code,d.name destination_name
      FROM inventory_movements m JOIN warehouses w ON w.id=m.warehouse_id
      LEFT JOIN warehouses d ON d.id=m.destination_warehouse_id
      WHERE m.item_id=:item ORDER BY m.movement_date DESC,m.id DESC""",{"item":item_id})
    return render_template("inventory_item_view.html",item=item,balances=balances,movements=movements)

@app.route("/inventory/movements",methods=["GET","POST"])
@login_required
def inventory_movements():
    if request.method=="POST":
        try:
            record_inventory_movement(
              request.form["movement_date"],request.form["movement_type"],
              int(request.form["item_id"]),int(request.form["warehouse_id"]),
              request.form["quantity"],request.form.get("unit_cost") or 0,
              reference_no=request.form.get("reference_no",""),
              batch_no=request.form.get("batch_no",""),serial_no=request.form.get("serial_no",""),
              production_date=request.form.get("production_date") or None,
              expiry_date=request.form.get("expiry_date") or None,
              notes=request.form.get("notes",""))
            flash("تم تسجيل حركة المخزون","success")
        except Exception as exc:
            flash(str(exc),"danger")
        return redirect(url_for("inventory_movements"))
    data=rows("""SELECT m.*,i.sku,i.name item_name,w.code warehouse_code,w.name warehouse_name
                 FROM inventory_movements m JOIN inventory i ON i.id=m.item_id
                 JOIN warehouses w ON w.id=m.warehouse_id
                 ORDER BY m.movement_date DESC,m.id DESC LIMIT 500""")
    return render_template("inventory_movements.html",movements=data,
      items=rows("SELECT id,sku,name,unit,cost FROM inventory WHERE active=1 ORDER BY name"),
      warehouses=rows("SELECT id,code,name FROM warehouses WHERE active=1 ORDER BY code"))

@app.route("/inventory/transfers",methods=["GET","POST"])
@login_required
def inventory_transfers():
    if request.method=="POST":
        item_id=int(request.form["item_id"]); source=int(request.form["source_warehouse_id"])
        destination=int(request.form["destination_warehouse_id"])
        if source==destination:
            flash("يجب اختيار مستودعين مختلفين","danger")
            return redirect(url_for("inventory_transfers"))
        qty=float(request.form["quantity"]); dt=request.form["transfer_date"]
        item=row("SELECT cost FROM inventory WHERE id=:id",{"id":item_id})
        try:
            ref=f"TR-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            record_inventory_movement(dt,"تحويل صادر",item_id,source,qty,item["cost"],
                                      destination_warehouse_id=destination,reference_type="TRANSFER",
                                      reference_no=ref,notes=request.form.get("notes",""))
            record_inventory_movement(dt,"تحويل وارد",item_id,destination,qty,item["cost"],
                                      destination_warehouse_id=source,reference_type="TRANSFER",
                                      reference_no=ref,notes=request.form.get("notes",""))
            flash("تم تحويل المخزون بين المستودعات","success")
        except Exception as exc:
            flash(str(exc),"danger")
        return redirect(url_for("inventory_transfers"))
    return render_template("inventory_transfers.html",
      items=rows("SELECT id,sku,name,unit FROM inventory WHERE active=1 ORDER BY name"),
      warehouses=rows("SELECT id,code,name FROM warehouses WHERE active=1 ORDER BY code"),
      transfers=rows("""SELECT m.*,i.name item_name,w.name source_name,d.name destination_name
        FROM inventory_movements m JOIN inventory i ON i.id=m.item_id
        JOIN warehouses w ON w.id=m.warehouse_id
        LEFT JOIN warehouses d ON d.id=m.destination_warehouse_id
        WHERE m.movement_type='تحويل صادر' ORDER BY m.id DESC LIMIT 200"""))

@app.route("/inventory/counts",methods=["GET","POST"])
@login_required
def inventory_counts():
    if request.method=="POST":
        count_date=request.form["count_date"]; warehouse_id=int(request.form["warehouse_id"])
        no=next_inventory_count_number(count_date)
        execute("""INSERT INTO inventory_counts(count_no,count_date,warehouse_id,status,notes,created_by,created_at)
          VALUES(:no,:dt,:warehouse,'معتمد',:notes,:uid,:created)""",
          {"no":no,"dt":count_date,"warehouse":warehouse_id,"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        count_id=row("SELECT id FROM inventory_counts WHERE count_no=:no",{"no":no})["id"]
        item_ids=request.form.getlist("item_id[]"); counted=request.form.getlist("counted_qty[]")
        for idx,item_id in enumerate(item_ids):
            item=row("SELECT cost FROM inventory WHERE id=:id",{"id":item_id})
            system=warehouse_stock(int(item_id),warehouse_id)
            actual=float(counted[idx] or 0); variance=round(actual-system,3)
            execute("""INSERT INTO inventory_count_items(count_id,item_id,system_qty,counted_qty,variance_qty,unit_cost)
                       VALUES(:count,:item,:system,:actual,:variance,:cost)""",
                    {"count":count_id,"item":item_id,"system":system,"actual":actual,
                     "variance":variance,"cost":item["cost"]})
            if variance>0:
                record_inventory_movement(count_date,"تسوية زيادة",int(item_id),warehouse_id,variance,item["cost"],
                                          reference_type="STOCK_COUNT",reference_id=count_id,reference_no=no)
            elif variance<0:
                record_inventory_movement(count_date,"تسوية نقص",int(item_id),warehouse_id,abs(variance),item["cost"],
                                          reference_type="STOCK_COUNT",reference_id=count_id,reference_no=no)
        flash(f"تم اعتماد الجرد {no} وترحيل الفروقات","success")
        return redirect(url_for("inventory_counts"))
    items=rows("SELECT id,sku,name,unit,cost FROM inventory WHERE active=1 ORDER BY name")
    return render_template("inventory_counts.html",
      items=items,warehouses=rows("SELECT id,code,name FROM warehouses WHERE active=1 ORDER BY code"),
      counts=rows("""SELECT c.*,w.code warehouse_code,w.name warehouse_name
                     FROM inventory_counts c JOIN warehouses w ON w.id=c.warehouse_id
                     ORDER BY c.count_date DESC,c.id DESC"""))

@app.route("/inventory/warehouses",methods=["GET","POST"])
@login_required
def warehouses():
    if request.method=="POST":
        execute("""INSERT INTO warehouses(code,name,branch_id,active)
                   VALUES(:code,:name,:branch,1) ON CONFLICT(code) DO NOTHING""",
                {"code":request.form["code"],"name":request.form["name"],
                 "branch":request.form.get("branch_id") or None})
        flash("تم حفظ المستودع","success")
        return redirect(url_for("warehouses"))
    return render_template("warehouses.html",
      warehouses=rows("""SELECT w.*,b.name branch_name FROM warehouses w
                         LEFT JOIN branches b ON b.id=w.branch_id ORDER BY w.code"""),
      branches=rows("SELECT id,name FROM branches WHERE active=1 ORDER BY name"))

@app.route("/inventory/export.xlsx")
@login_required
def inventory_export():
    data=rows("""SELECT sku,name,name_en,barcode,unit,quantity,cost,sale_price,reorder_level,
                 min_level,max_level,valuation_method,active FROM inventory ORDER BY sku,name""")
    headers=["الكود","اسم الصنف","English Name","الباركود","الوحدة","الرصيد","متوسط التكلفة",
             "سعر البيع","إعادة الطلب","الحد الأدنى","الحد الأعلى","طريقة التقييم","الحالة"]
    keys=["sku","name","name_en","barcode","unit","quantity","cost","sale_price","reorder_level",
          "min_level","max_level","valuation_method","active"]
    return xlsx_response("inventory.xlsx","المخزون",headers,[[r.get(k) for k in keys] for r in data])


@app.route("/employees", methods=["GET","POST"])
@login_required
def employees():
    if request.method == "POST":
        execute("""INSERT INTO employees(employee_no,name,job_title,branch_id,basic_salary,allowances)
                   VALUES(:eno,:n,:job,:bid,:sal,:allow)""",
                {"eno": request.form["employee_no"] or None, "n": request.form["name"],
                 "job": request.form["job_title"], "bid": request.form["branch_id"],
                 "sal": float(request.form["basic_salary"] or 0),
                 "allow": float(request.form["allowances"] or 0)})
        audit("CREATE", "EMPLOYEE", f"إضافة موظف: {request.form['name']}")
        flash("تمت إضافة الموظف", "success")
    employee_rows = rows("""SELECT e.*, b.name branch_name FROM employees e
                            LEFT JOIN branches b ON b.id=e.branch_id ORDER BY e.id DESC""")
    return render_template("employees.html", rows=employee_rows,
                           branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"))



def zatca_tlv_base64(seller_name, vat_number, invoice_datetime, total, vat_total):
    """
    Create the QR payload using only the five standard TLV fields:
    1 seller name
    2 VAT number
    3 invoice date/time
    4 invoice total including VAT
    5 VAT total

    The invoice number and UUID are intentionally NOT included in the QR payload.
    """
    values = [
        (1, str(seller_name or "")),
        (2, str(vat_number or "")),
        (3, str(invoice_datetime or "")),
        (4, f"{Decimal(str(total)):.2f}"),
        (5, f"{Decimal(str(vat_total)):.2f}"),
    ]

    payload = bytearray()
    for tag, value in values:
        encoded = value.encode("utf-8")
        if len(encoded) > 255:
            raise ValueError("QR field is too long")
        payload.extend(bytes([tag, len(encoded)]))
        payload.extend(encoded)

    return base64.b64encode(payload).decode("ascii")


def qr_data_uri(data):
    qr = qrcode.QRCode(version=None, box_size=7, border=2)
    qr.add_data(data)
    qr.make(fit=True)
    image = qr.make_image(fill_color="black", back_color="white")
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")


@app.route("/invoice/<int:invoice_id>")
@login_required
def invoice_view(invoice_id):
    invoice = row("""
        SELECT i.*, c.name customer_name, c.name_en customer_name_en, c.vat_number customer_vat,
               c.phone customer_phone, c.email customer_email,
               b.name branch_name
        FROM invoices i
        JOIN customers c ON c.id=i.customer_id
        LEFT JOIN branches b ON b.id=i.branch_id
        WHERE i.id=:id
    """, {"id": invoice_id})

    if not invoice:
        return "الفاتورة غير موجودة", 404

    items = rows("""
        SELECT *
        FROM invoice_items
        WHERE invoice_id=:invoice_id
        ORDER BY id
    """, {"invoice_id": invoice_id})

    company = row("SELECT * FROM settings WHERE id=1")
    issue_time = invoice["created_at"].isoformat() if invoice["created_at"] else str(invoice["invoice_date"])

    tlv = zatca_tlv_base64(
        company["company_name_ar"],
        company["vat_number"],
        issue_time,
        invoice["total"],
        invoice["vat"]
    )
    qr_image = qr_data_uri(tlv)

    return render_template(
        "invoice_print.html",
        invoice=invoice,
        items=items,
        company=company,
        qr_image=qr_image,
        qr_payload=tlv,
        print_lang=request.args.get("lang","ar") if request.args.get("lang","ar") in ("ar","en","bi") else "ar"
    )



@app.route("/chart-of-accounts", methods=["GET","POST"])
@login_required
def chart_of_accounts():
    if request.method == "POST":
        parent_id = request.form.get("parent_id") or None
        parent_level = 0
        if parent_id:
            parent = row("SELECT level FROM chart_of_accounts WHERE id=:id", {"id": parent_id})
            parent_level = parent["level"] if parent else 0

        execute("""INSERT INTO chart_of_accounts(
            account_code,account_name_ar,account_name_en,account_type,
            parent_id,level,accepts_entries,normal_balance,statement_type,active)
            VALUES(:code,:ar,:en,:type,:parent,:level,:accepts,:normal_balance,:statement_type,:active)""",
            {"code":request.form["account_code"].strip(),
             "ar":request.form["account_name_ar"].strip(),
             "en":request.form.get("account_name_en","").strip(),
             "type":request.form["account_type"],
             "parent":parent_id,
             "level":parent_level+1,
             "accepts":1 if request.form.get("accepts_entries")=="1" else 0,
             "normal_balance":request.form.get("normal_balance","مدين"),
             "statement_type":request.form.get("statement_type","الميزانية العمومية"),
             "active":1 if request.form.get("active")=="1" else 0})
        audit("CREATE", "ACCOUNT", f"إضافة حساب {request.form['account_code']}")
        flash("تمت إضافة الحساب", "success")
        return redirect(url_for("chart_of_accounts"))

    q = request.args.get("q", "").strip()
    account_type = request.args.get("account_type", "").strip()
    active = request.args.get("active", "").strip()

    conditions = []
    params = {}
    if q:
        conditions.append("(a.account_code ILIKE :q OR a.account_name_ar ILIKE :q OR a.account_name_en ILIKE :q)")
        params["q"] = f"%{q}%"
    if account_type:
        conditions.append("a.account_type=:account_type")
        params["account_type"] = account_type
    if active in ("0","1"):
        conditions.append("a.active=:active")
        params["active"] = int(active)

    where = " WHERE " + " AND ".join(conditions) if conditions else ""

    accounts = rows(f"""
        SELECT a.*, p.account_name_ar parent_name,
               COALESCE(ch.children_count,0) children_count,
               COALESCE(mv.entries_count,0) entries_count,
               COALESCE(mv.total_debit,0) total_debit,
               COALESCE(mv.total_credit,0) total_credit,
               COALESCE(mv.total_debit,0)-COALESCE(mv.total_credit,0) balance
        FROM chart_of_accounts a
        LEFT JOIN chart_of_accounts p ON p.id=a.parent_id
        LEFT JOIN (
            SELECT parent_id, COUNT(*) children_count
            FROM chart_of_accounts
            WHERE parent_id IS NOT NULL
            GROUP BY parent_id
        ) ch ON ch.parent_id=a.id
        LEFT JOIN (
            SELECT account_id, COUNT(*) entries_count,
                   SUM(debit) total_debit, SUM(credit) total_credit
            FROM journal_entry_lines
            GROUP BY account_id
        ) mv ON mv.account_id=a.id
        {where}
        ORDER BY a.account_code
    """, params)

    parents = rows("""
        SELECT id,account_code,account_name_ar,level
        FROM chart_of_accounts
        ORDER BY account_code
    """)

    return render_template(
        "chart_of_accounts.html",
        accounts=accounts,
        parents=parents,
        q=q,
        selected_type=account_type,
        selected_active=active
    )



@app.route("/chart-of-accounts/next-code")
@login_required
def chart_next_code():
    parent_id = request.args.get("parent_id", type=int)
    if not parent_id:
        return {"code": "", "message": "اختر الحساب الرئيسي أولاً"}

    parent = row("""
        SELECT id, account_code, level
        FROM chart_of_accounts
        WHERE id=:id
    """, {"id": parent_id})

    if not parent:
        return {"code": "", "message": "الحساب الرئيسي غير موجود"}, 404

    children = rows("""
        SELECT account_code
        FROM chart_of_accounts
        WHERE parent_id=:id
        ORDER BY account_code
    """, {"id": parent_id})

    parent_code = str(parent["account_code"]).strip()
    numeric_codes = []

    for child in children:
        code = str(child["account_code"]).strip()
        if code.isdigit():
            numeric_codes.append(int(code))

    if parent_code.isdigit():
        parent_number = int(parent_code)
        parent_level = int(parent["level"] or 1)

        # 1000 -> 1100, 1200 ...
        # 1100 -> 1110, 1120 ...
        # 1110 -> 1111, 1112 ...
        step = 10 ** max(0, 3 - parent_level)

        if numeric_codes:
            candidate = max(numeric_codes) + step
        else:
            candidate = parent_number + step

        width = max(len(parent_code), len(str(candidate)))
        suggested_code = str(candidate).zfill(width)
    else:
        # Alphanumeric fallback: ABC -> ABC-01, ABC-02 ...
        sequence = 1
        existing = {str(c["account_code"]).strip() for c in children}
        while f"{parent_code}-{sequence:02d}" in existing:
            sequence += 1
        suggested_code = f"{parent_code}-{sequence:02d}"

    return {
        "code": suggested_code,
        "parent_code": parent_code,
        "message": "تم اقتراح الكود تلقائيًا"
    }


@app.route("/chart-of-accounts/<int:account_id>")
@login_required
def account_view(account_id):
    account = row("""
        SELECT a.*,p.account_name_ar parent_name,
               COALESCE(mv.entries_count,0) entries_count,
               COALESCE(mv.total_debit,0) total_debit,
               COALESCE(mv.total_credit,0) total_credit,
               COALESCE(mv.total_debit,0)-COALESCE(mv.total_credit,0) balance
        FROM chart_of_accounts a
        LEFT JOIN chart_of_accounts p ON p.id=a.parent_id
        LEFT JOIN (
            SELECT account_id,COUNT(*) entries_count,
                   SUM(debit) total_debit,SUM(credit) total_credit
            FROM journal_entry_lines GROUP BY account_id
        ) mv ON mv.account_id=a.id
        WHERE a.id=:id
    """, {"id": account_id})
    if not account:
        return "الحساب غير موجود", 404

    children = rows("""
        SELECT id,account_code,account_name_ar,accepts_entries,active
        FROM chart_of_accounts WHERE parent_id=:id ORDER BY account_code
    """, {"id": account_id})

    movements = rows("""
        SELECT j.journal_no,j.journal_date,j.status,l.debit,l.credit,
               l.line_description,l.invoice_number,l.invoice_date
        FROM journal_entry_lines l
        JOIN journal_entries j ON j.id=l.journal_id
        WHERE l.account_id=:id
        ORDER BY j.journal_date DESC,j.id DESC,l.id DESC
        LIMIT 200
    """, {"id": account_id})

    company = row("SELECT * FROM settings WHERE id=1")
    return render_template("account_view.html", account=account, children=children,
                           movements=movements, company=company)

@app.route("/chart-of-accounts/<int:account_id>/edit", methods=["GET","POST"])
@login_required
def account_edit(account_id):
    account = row("SELECT * FROM chart_of_accounts WHERE id=:id", {"id": account_id})
    if not account:
        return "الحساب غير موجود", 404

    if request.method == "POST":
        parent_id = request.form.get("parent_id") or None
        if parent_id and int(parent_id) == account_id:
            flash("لا يمكن جعل الحساب أبًا لنفسه", "danger")
            return redirect(url_for("account_edit", account_id=account_id))

        parent_level = 0
        if parent_id:
            parent = row("SELECT level FROM chart_of_accounts WHERE id=:id", {"id": parent_id})
            parent_level = parent["level"] if parent else 0

        execute("""UPDATE chart_of_accounts SET
            account_code=:code,account_name_ar=:ar,account_name_en=:en,
            account_type=:type,parent_id=:parent,level=:level,
            accepts_entries=:accepts,normal_balance=:normal_balance,
            statement_type=:statement_type,active=:active
            WHERE id=:id""",
            {"code":request.form["account_code"].strip(),
             "ar":request.form["account_name_ar"].strip(),
             "en":request.form.get("account_name_en","").strip(),
             "type":request.form["account_type"],
             "parent":parent_id,
             "level":parent_level+1,
             "accepts":1 if request.form.get("accepts_entries")=="1" else 0,
             "normal_balance":request.form["normal_balance"],
             "statement_type":request.form["statement_type"],
             "active":1 if request.form.get("active")=="1" else 0,
             "id":account_id})
        audit("UPDATE", "ACCOUNT", f"تعديل الحساب {request.form['account_code']}")
        flash("تم تعديل الحساب", "success")
        return redirect(url_for("account_view", account_id=account_id))

    parents = rows("""
        SELECT id,account_code,account_name_ar,level
        FROM chart_of_accounts WHERE id<>:id ORDER BY account_code
    """, {"id": account_id})
    return render_template("account_edit.html", account=account, parents=parents)

@app.route("/chart-of-accounts/<int:account_id>/delete", methods=["POST"])
@login_required
def account_delete(account_id):
    account = row("SELECT * FROM chart_of_accounts WHERE id=:id", {"id": account_id})
    if not account:
        return "الحساب غير موجود", 404

    children_count = row("SELECT COUNT(*) count FROM chart_of_accounts WHERE parent_id=:id", {"id": account_id})["count"]
    entries_count = row("SELECT COUNT(*) count FROM journal_entry_lines WHERE account_id=:id", {"id": account_id})["count"]

    if children_count > 0:
        flash("لا يمكن حذف الحساب لأنه يحتوي على حسابات فرعية", "danger")
    elif entries_count > 0:
        flash("لا يمكن حذف الحساب لأنه مستخدم في قيود يومية. يمكنك تعطيله بدلًا من حذفه.", "danger")
    else:
        execute("DELETE FROM chart_of_accounts WHERE id=:id", {"id": account_id})
        audit("DELETE", "ACCOUNT", f"حذف الحساب {account['account_code']}")
        flash("تم حذف الحساب", "success")
    return redirect(url_for("chart_of_accounts"))

@app.route("/chart-of-accounts/export.xlsx")
@login_required
def chart_export():
    data = rows("""
        SELECT a.account_code,a.account_name_ar,a.account_name_en,a.account_type,
               p.account_code parent_code,p.account_name_ar parent_name,
               a.level,a.accepts_entries,a.normal_balance,a.statement_type,a.active,
               COALESCE(mv.entries_count,0) entries_count,
               COALESCE(mv.total_debit,0) total_debit,
               COALESCE(mv.total_credit,0) total_credit,
               COALESCE(mv.total_debit,0)-COALESCE(mv.total_credit,0) balance
        FROM chart_of_accounts a
        LEFT JOIN chart_of_accounts p ON p.id=a.parent_id
        LEFT JOIN (
            SELECT account_id,COUNT(*) entries_count,
                   SUM(debit) total_debit,SUM(credit) total_credit
            FROM journal_entry_lines GROUP BY account_id
        ) mv ON mv.account_id=a.id
        ORDER BY a.account_code
    """)
    headers = ["الكود","اسم الحساب","الاسم الإنجليزي","النوع","كود الحساب الأب",
               "الحساب الأب","المستوى","يقبل حركة","طبيعة الرصيد","القائمة المالية",
               "الحالة","عدد القيود","إجمالي المدين","إجمالي الدائن","الرصيد"]
    keys = ["account_code","account_name_ar","account_name_en","account_type",
            "parent_code","parent_name","level","accepts_entries","normal_balance",
            "statement_type","active","entries_count","total_debit","total_credit","balance"]
    records = [[r.get(k) for k in keys] for r in data]
    return xlsx_response("chart_of_accounts.xlsx","دليل الحسابات",headers,records)



@app.route("/cost-centers", methods=["GET","POST"])
@login_required
def cost_centers():
    if request.method=="POST":
        execute("INSERT INTO cost_centers(code,name,active) VALUES(:c,:n,1)",
                {"c":request.form["code"].strip(),"n":request.form["name"].strip()})
        flash("تمت إضافة مركز التكلفة","success")
        return redirect(url_for("cost_centers"))
    return render_template("cost_centers.html",
        centers=rows("SELECT * FROM cost_centers ORDER BY code"))

@app.route("/journal-entries", methods=["GET","POST"])
@login_required
def journal_entries():
    if request.method=="POST":
        account_ids=request.form.getlist("account_id[]")
        debits=request.form.getlist("debit[]")
        credits=request.form.getlist("credit[]")
        taxable_values=request.form.getlist("taxable[]")
        tax_directions=request.form.getlist("tax_direction[]")
        supplier_ids=request.form.getlist("supplier_id[]")
        customer_ids=request.form.getlist("customer_id[]")
        party_types=request.form.getlist("party_type[]")
        tax_numbers=request.form.getlist("tax_number[]")
        invoice_numbers=request.form.getlist("invoice_number[]")
        invoice_dates=request.form.getlist("invoice_date[]")
        descriptions=request.form.getlist("line_description[]")
        cost_center_ids=request.form.getlist("cost_center_id[]")
        lines=[]; total_debit=0.0; total_credit=0.0
        for i,account_id in enumerate(account_ids):
            if not account_id: continue
            debit=float(debits[i] or 0); credit=float(credits[i] or 0)
            if debit>0 and credit>0:
                flash(f"السطر {i+1}: لا يمكن إدخال مدين ودائن معًا","danger")
                return redirect(url_for("journal_entries"))
            if debit<=0 and credit<=0:
                flash(f"السطر {i+1}: أدخل مبلغ مدين أو دائن","danger")
                return redirect(url_for("journal_entries"))
            account=row("SELECT accepts_entries FROM chart_of_accounts WHERE id=:id",{"id":account_id})
            if not account or account["accepts_entries"]!=1:
                flash(f"السطر {i+1}: الحساب لا يقبل حركة","danger")
                return redirect(url_for("journal_entries"))
            taxable=1 if i<len(taxable_values) and taxable_values[i]=="1" else 0
            direction=tax_directions[i] if i<len(tax_directions) else "غير مطبق"
            party_type = party_types[i] if i < len(party_types) else ""
            supplier_id = (supplier_ids[i] or None) if i < len(supplier_ids) else None
            customer_id = (customer_ids[i] or None) if i < len(customer_ids) else None
            tax_number = tax_numbers[i].strip() if i < len(tax_numbers) else ""
            invoice_number = invoice_numbers[i].strip() if i < len(invoice_numbers) else ""
            invoice_date = invoice_dates[i] if i < len(invoice_dates) and invoice_dates[i] else None

            if party_type == "مورد":
                customer_id = None
                party = row("SELECT vat_number FROM suppliers WHERE id=:id", {"id": supplier_id}) if supplier_id else None
                tax_number = (party["vat_number"] or "") if party else ""
            elif party_type == "عميل":
                supplier_id = None
                party = row("SELECT vat_number FROM customers WHERE id=:id", {"id": customer_id}) if customer_id else None
                tax_number = (party["vat_number"] or "") if party else ""
            else:
                supplier_id = None
                customer_id = None
                tax_number = ""

            if taxable:
                if direction == "غير مطبق":
                    flash(f"السطر {i+1}: حدد مدخلات أو مخرجات","danger")
                    return redirect(url_for("journal_entries"))
                if party_type not in ("مورد", "عميل"):
                    flash(f"السطر {i+1}: حدد نوع الطرف مورد أو عميل","danger")
                    return redirect(url_for("journal_entries"))
                if party_type == "مورد" and not supplier_id:
                    flash(f"السطر {i+1}: اختر المورد","danger")
                    return redirect(url_for("journal_entries"))
                if party_type == "عميل" and not customer_id:
                    flash(f"السطر {i+1}: اختر العميل","danger")
                    return redirect(url_for("journal_entries"))
                if not tax_number:
                    flash(f"السطر {i+1}: الرقم الضريبي غير موجود في بيانات الطرف","danger")
                    return redirect(url_for("journal_entries"))
                if not invoice_number:
                    flash(f"السطر {i+1}: رقم الفاتورة مطلوب للعملية الخاضعة للضريبة","danger")
                    return redirect(url_for("journal_entries"))
                if not invoice_date:
                    flash(f"السطر {i+1}: تاريخ الفاتورة مطلوب للعملية الخاضعة للضريبة","danger")
                    return redirect(url_for("journal_entries"))

            lines.append({"account_id":account_id,"debit":round(debit,2),"credit":round(credit,2),
                          "taxable":taxable,"tax_direction":direction,
                          "supplier_id":supplier_id,
                          "customer_id":customer_id,
                          "party_type":party_type,
                          "tax_number":tax_number,
                          "invoice_number":invoice_number,
                          "invoice_date":invoice_date,
                          "line_description":descriptions[i] if i<len(descriptions) else "",
                          "cost_center_id":(cost_center_ids[i] or None) if i<len(cost_center_ids) else None})
            total_debit+=debit; total_credit+=credit
        if len(lines)<2:
            flash("يجب أن يحتوي القيد على سطرين على الأقل","danger")
            return redirect(url_for("journal_entries"))
        total_debit=round(total_debit,2); total_credit=round(total_credit,2)
        if total_debit!=total_credit:
            flash(f"القيد غير متوازن. الفرق {abs(total_debit-total_credit):.2f}","danger")
            return redirect(url_for("journal_entries"))
        try:
            ensure_open_period(request.form["journal_date"])
        except Exception as exc:
            flash(str(exc),"danger")
            return redirect(url_for("journal_entries"))
        journal_no=next_journal_number(request.form["journal_date"])
        execute("""INSERT INTO journal_entries(
            journal_no,journal_date,reference,description,status,total_debit,total_credit,created_by,created_at)
            VALUES(:no,:date,:ref,:desc,:status,:debit,:credit,:user,:created)""",
            {"no":journal_no,"date":request.form["journal_date"],
             "ref":request.form.get("reference",""),"desc":request.form.get("description",""),
             "status":request.form["status"],"debit":total_debit,"credit":total_credit,
             "user":session.get("user_id"),"created":datetime.now()})
        journal_id=row("SELECT id FROM journal_entries WHERE journal_no=:no",{"no":journal_no})["id"]
        for line in lines:
            execute("""INSERT INTO journal_entry_lines(
                journal_id,account_id,debit,credit,taxable,tax_direction,supplier_id,customer_id,
                party_type,tax_number,invoice_number,invoice_date,line_description,cost_center_id)
                VALUES(:journal_id,:account_id,:debit,:credit,:taxable,:tax_direction,
                :supplier_id,:customer_id,:party_type,:tax_number,:invoice_number,:invoice_date,
                :line_description,:cost_center_id)""",
                {"journal_id":journal_id,**line})
        flash(f"تم حفظ القيد {journal_no}","success")
        if request.form.get("print_after_save") == "1":
            return redirect(url_for("journal_view", journal_id=journal_id, print=1))
        return redirect(url_for("journal_entries"))
    q = request.args.get("q", "").strip()
    journal_params = {}
    journal_where = ""
    if q:
        journal_where = " WHERE journal_no ILIKE :q OR description ILIKE :q OR reference ILIKE :q "
        journal_params["q"] = f"%{q}%"
    return render_template("journal_entries.html",
        journals=rows(f"SELECT * FROM journal_entries {journal_where} ORDER BY journal_date DESC,id DESC", journal_params),
        q=q,
        accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                         WHERE active=1 AND accepts_entries=1 ORDER BY account_code"""),
        suppliers=rows("SELECT id,name,vat_number FROM suppliers ORDER BY name"),
        customers=rows("SELECT id,name,vat_number FROM customers ORDER BY name"),
        centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"))



# ==================== Complete CRUD / View / Print / Export ====================

MASTER_MODULES = {
    "customers": {
        "table": "customers", "title": "العملاء", "entity": "CUSTOMER",
        "fields": ["name", "vat_number", "phone", "email"],
        "labels": ["الاسم", "الرقم الضريبي", "الهاتف", "البريد الإلكتروني"],
    },
    "suppliers": {
        "table": "suppliers", "title": "الموردون", "entity": "SUPPLIER",
        "fields": ["name", "vat_number", "phone", "email"],
        "labels": ["الاسم", "الرقم الضريبي", "الهاتف", "البريد الإلكتروني"],
    },
    "inventory": {
        "table": "inventory", "title": "المخزون", "entity": "INVENTORY",
        "fields": ["sku", "name", "quantity", "unit", "cost", "sale_price", "reorder_level"],
        "labels": ["الكود", "اسم الصنف", "الكمية", "الوحدة", "التكلفة", "سعر البيع", "حد إعادة الطلب"],
    },
    "employees": {
        "table": "employees", "title": "الموظفون", "entity": "EMPLOYEE",
        "fields": ["employee_no", "name", "job_title", "basic_salary", "allowances", "active"],
        "labels": ["رقم الموظف", "الاسم", "المسمى الوظيفي", "الراتب الأساسي", "البدلات", "الحالة"],
    },
    "branches": {
        "table": "branches", "title": "الفروع", "entity": "BRANCH",
        "fields": ["name", "city", "active"],
        "labels": ["اسم الفرع", "المدينة", "الحالة"],
    },
    "cost-centers": {
        "table": "cost_centers", "title": "مراكز التكلفة", "entity": "COST_CENTER",
        "fields": ["code", "name", "active"],
        "labels": ["الكود", "اسم مركز التكلفة", "الحالة"],
    },
}

@app.route("/manage/<module_name>")
@login_required
def manage_module(module_name):
    cfg = MASTER_MODULES.get(module_name)
    if not cfg:
        return "الوحدة غير موجودة", 404
    q = request.args.get("q", "").strip()
    where = ""
    params = {}
    if q:
        searchable = [f"CAST({f} AS TEXT) ILIKE :q" for f in cfg["fields"]]
        where = " WHERE " + " OR ".join(searchable)
        params["q"] = f"%{q}%"
    data = rows(f"SELECT * FROM {cfg['table']}{where} ORDER BY id DESC", params)
    return render_template("module_manage.html", module_name=module_name, cfg=cfg, data=data, q=q)

@app.route("/manage/<module_name>/<int:record_id>")
@login_required
def module_view(module_name, record_id):
    cfg = MASTER_MODULES.get(module_name)
    if not cfg:
        return "الوحدة غير موجودة", 404
    record = row(f"SELECT * FROM {cfg['table']} WHERE id=:id", {"id": record_id})
    if not record:
        return "السجل غير موجود", 404
    return render_template("module_view.html", module_name=module_name, cfg=cfg, record=record)

@app.route("/manage/<module_name>/<int:record_id>/edit", methods=["GET", "POST"])
@login_required
def module_edit(module_name, record_id):
    cfg = MASTER_MODULES.get(module_name)
    if not cfg:
        return "الوحدة غير موجودة", 404
    record = row(f"SELECT * FROM {cfg['table']} WHERE id=:id", {"id": record_id})
    if not record:
        return "السجل غير موجود", 404
    if request.method == "POST":
        assignments = ", ".join([f"{f}=:{f}" for f in cfg["fields"]])
        params = {"id": record_id}
        for field in cfg["fields"]:
            value = request.form.get(field, "")
            if field in ("quantity", "cost", "sale_price", "reorder_level", "basic_salary", "allowances"):
                value = float(value or 0)
            elif field == "active":
                value = 1 if value == "1" else 0
            params[field] = value
        execute(f"UPDATE {cfg['table']} SET {assignments} WHERE id=:id", params)
        audit("UPDATE", cfg["entity"], f"تعديل السجل رقم {record_id}")
        flash("تم حفظ التعديل", "success")
        return redirect(url_for("module_view", module_name=module_name, record_id=record_id))
    return render_template("module_edit.html", module_name=module_name, cfg=cfg, record=record)

@app.route("/manage/<module_name>/<int:record_id>/delete", methods=["POST"])
@login_required
def module_delete(module_name, record_id):
    cfg = MASTER_MODULES.get(module_name)
    if not cfg:
        return "الوحدة غير موجودة", 404
    try:
        execute(f"DELETE FROM {cfg['table']} WHERE id=:id", {"id": record_id})
        audit("DELETE", cfg["entity"], f"حذف السجل رقم {record_id}")
        flash("تم حذف السجل", "success")
    except Exception:
        db.session.rollback()
        flash("تعذر الحذف لأن السجل مرتبط بعمليات أخرى. يمكن تعطيله بدلًا من حذفه.", "danger")
    return redirect(url_for("manage_module", module_name=module_name))

@app.route("/manage/<module_name>/export.xlsx")
@login_required
def module_export(module_name):
    cfg = MASTER_MODULES.get(module_name)
    if not cfg:
        return "الوحدة غير موجودة", 404
    data = rows(f"SELECT * FROM {cfg['table']} ORDER BY id")
    records = [[r.get(f) for f in cfg["fields"]] for r in data]
    return xlsx_response(f"{module_name}.xlsx", cfg["title"], cfg["labels"], records)

@app.route("/journal-entries/<int:journal_id>")
@login_required
def journal_view(journal_id):
    journal = row("SELECT * FROM journal_entries WHERE id=:id", {"id": journal_id})
    if not journal:
        return "القيد غير موجود", 404
    lines = rows("""
        SELECT l.*, a.account_code, a.account_name_ar,
               s.name supplier_name, cu.name customer_name,
               c.code cost_center_code, c.name cost_center_name
        FROM journal_entry_lines l
        JOIN chart_of_accounts a ON a.id=l.account_id
        LEFT JOIN suppliers s ON s.id=l.supplier_id
        LEFT JOIN customers cu ON cu.id=l.customer_id
        LEFT JOIN cost_centers c ON c.id=l.cost_center_id
        WHERE l.journal_id=:id ORDER BY l.id
    """, {"id": journal_id})
    company = row("SELECT * FROM settings WHERE id=1")
    return render_template("journal_view.html", journal=journal, lines=lines, company=company)

@app.route("/journal-entries/<int:journal_id>/delete", methods=["POST"])
@login_required
def journal_delete(journal_id):
    journal = row("SELECT * FROM journal_entries WHERE id=:id", {"id": journal_id})
    if not journal:
        return "القيد غير موجود", 404
    if journal["status"] == "مرحّل":
        flash("لا يمكن حذف قيد مرحّل. يجب إنشاء قيد عكسي.", "danger")
    else:
        execute("DELETE FROM journal_entries WHERE id=:id", {"id": journal_id})
        audit("DELETE", "JOURNAL", f"حذف القيد {journal['journal_no']}")
        flash("تم حذف القيد المسودة", "success")
    return redirect(url_for("journal_entries"))

@app.route("/journal-entries/export.xlsx")
@login_required
def journal_export():
    q = request.args.get("q", "").strip()
    params = {}
    where = ""
    if q:
        where = """ WHERE j.journal_no ILIKE :q OR j.description ILIKE :q
                    OR a.account_code ILIKE :q OR a.account_name_ar ILIKE :q
                    OR COALESCE(l.invoice_number,'') ILIKE :q """
        params["q"] = f"%{q}%"
    data = rows(f"""
        SELECT j.journal_no,j.journal_date,j.status,j.reference,j.description,
               a.account_code,a.account_name_ar,l.debit,l.credit,l.taxable,
               l.tax_direction,l.party_type,s.name supplier_name,cu.name customer_name,
               l.tax_number,l.invoice_number,l.invoice_date,
               l.line_description,c.code cost_center_code,c.name cost_center_name
        FROM journal_entries j
        JOIN journal_entry_lines l ON l.journal_id=j.id
        JOIN chart_of_accounts a ON a.id=l.account_id
        LEFT JOIN suppliers s ON s.id=l.supplier_id
        LEFT JOIN customers cu ON cu.id=l.customer_id
        LEFT JOIN cost_centers c ON c.id=l.cost_center_id
        {where}
        ORDER BY j.journal_date DESC,j.id DESC,l.id
    """, params)
    headers = ["رقم القيد","تاريخ القيد","الحالة","المرجع","البيان العام","الكود","اسم الحساب",
               "مدين","دائن","خاضع للضريبة","مدخلات/مخرجات","نوع الطرف","المورد","العميل","الرقم الضريبي","رقم الفاتورة",
               "تاريخ الفاتورة","البيان","كود مركز التكلفة","مركز التكلفة"]
    records = [[r.get(k) for k in [
        "journal_no","journal_date","status","reference","description","account_code",
        "account_name_ar","debit","credit","taxable","tax_direction","party_type","supplier_name",
        "customer_name","tax_number","invoice_number","invoice_date","line_description",
        "cost_center_code","cost_center_name"
    ]] for r in data]
    return xlsx_response("journal_entries.xlsx", "قيود اليومية", headers, records)

@app.route("/invoices/<int:invoice_id>/delete", methods=["POST"])
@login_required
def invoice_delete(invoice_id):
    invoice = row("SELECT * FROM invoices WHERE id=:id", {"id": invoice_id})
    if not invoice:
        return "الفاتورة غير موجودة", 404
    if invoice["status"] == "معتمدة":
        flash("لا يمكن حذف فاتورة معتمدة. غيّر حالتها أو أنشئ إشعارًا دائنًا.", "danger")
    else:
        execute("DELETE FROM invoices WHERE id=:id", {"id": invoice_id})
        audit("DELETE", "INVOICE", f"حذف الفاتورة {invoice['invoice_no']}")
        flash("تم حذف الفاتورة المسودة", "success")
    return redirect(url_for("invoices"))




@app.route("/multi-journal", methods=["GET", "POST"])
@login_required
def multi_journal():
    if request.method == "POST":
        group_numbers = request.form.getlist("group_no[]")
        account_ids = request.form.getlist("account_id[]")
        debits = request.form.getlist("debit[]")
        credits = request.form.getlist("credit[]")
        taxable_values = request.form.getlist("taxable[]")
        tax_directions = request.form.getlist("tax_direction[]")
        party_types = request.form.getlist("party_type[]")
        supplier_ids = request.form.getlist("supplier_id[]")
        customer_ids = request.form.getlist("customer_id[]")
        tax_numbers = request.form.getlist("tax_number[]")
        invoice_numbers = request.form.getlist("invoice_number[]")
        invoice_dates = request.form.getlist("invoice_date[]")
        descriptions = request.form.getlist("line_description[]")
        cost_center_ids = request.form.getlist("cost_center_id[]")

        groups = {}
        for i, account_id in enumerate(account_ids):
            if not account_id:
                continue
            group_no = int(group_numbers[i] or 1)
            debit = float(debits[i] or 0)
            credit = float(credits[i] or 0)
            if debit > 0 and credit > 0:
                flash(f"المجموعة {group_no}، السطر {i+1}: لا يمكن إدخال مدين ودائن معًا", "danger")
                return redirect(url_for("multi_journal"))
            if debit <= 0 and credit <= 0:
                continue

            taxable = 1 if taxable_values[i] == "1" else 0
            party_type = party_types[i] if i < len(party_types) else ""
            supplier_id = (supplier_ids[i] or None) if i < len(supplier_ids) else None
            customer_id = (customer_ids[i] or None) if i < len(customer_ids) else None
            tax_number = tax_numbers[i].strip() if i < len(tax_numbers) else ""
            direction = tax_directions[i] if i < len(tax_directions) else "غير مطبق"

            if taxable:
                if direction == "غير مطبق":
                    flash(f"المجموعة {group_no}: حدد مدخلات أو مخرجات", "danger")
                    return redirect(url_for("multi_journal"))
                if party_type not in ("مورد", "عميل"):
                    flash(f"المجموعة {group_no}: حدد موردًا أو عميلًا", "danger")
                    return redirect(url_for("multi_journal"))
                if not tax_number:
                    flash(f"المجموعة {group_no}: الرقم الضريبي مطلوب", "danger")
                    return redirect(url_for("multi_journal"))
                if not invoice_numbers[i] or not invoice_dates[i]:
                    flash(f"المجموعة {group_no}: رقم وتاريخ الفاتورة مطلوبان", "danger")
                    return redirect(url_for("multi_journal"))

            groups.setdefault(group_no, []).append({
                "account_id": account_id,
                "debit": round(debit, 2),
                "credit": round(credit, 2),
                "taxable": taxable,
                "tax_direction": direction,
                "party_type": party_type,
                "supplier_id": supplier_id if party_type == "مورد" else None,
                "customer_id": customer_id if party_type == "عميل" else None,
                "tax_number": tax_number,
                "invoice_number": invoice_numbers[i] if i < len(invoice_numbers) else "",
                "invoice_date": invoice_dates[i] if i < len(invoice_dates) and invoice_dates[i] else None,
                "line_description": descriptions[i] if i < len(descriptions) else "",
                "cost_center_id": (cost_center_ids[i] or None) if i < len(cost_center_ids) else None,
            })

        if not groups:
            flash("أدخل مجموعة واحدة على الأقل", "danger")
            return redirect(url_for("multi_journal"))

        for group_no, lines in groups.items():
            td = round(sum(x["debit"] for x in lines), 2)
            tc = round(sum(x["credit"] for x in lines), 2)
            if td != tc:
                flash(f"المجموعة {group_no} غير متوازنة، الفرق {abs(td-tc):.2f}", "danger")
                return redirect(url_for("multi_journal"))

        batch_no = next_batch_number(request.form["batch_date"])
        execute("""INSERT INTO journal_batches(batch_no,batch_date,description,status,created_by,created_at)
                   VALUES(:no,:date,:description,:status,:user,:created)""",
                {"no": batch_no, "date": request.form["batch_date"],
                 "description": request.form.get("description",""),
                 "status": request.form["status"],
                 "user": session.get("user_id"), "created": datetime.now()})
        batch_id = row("SELECT id FROM journal_batches WHERE batch_no=:no", {"no": batch_no})["id"]

        for group_no, lines in sorted(groups.items()):
            td = round(sum(x["debit"] for x in lines), 2)
            tc = round(sum(x["credit"] for x in lines), 2)
            journal_no = next_journal_number(request.form["batch_date"])
            execute("""INSERT INTO journal_batch_groups(batch_id,group_no,journal_no,total_debit,total_credit,status)
                       VALUES(:batch,:group_no,:journal_no,:td,:tc,:status)""",
                    {"batch": batch_id, "group_no": group_no, "journal_no": journal_no,
                     "td": td, "tc": tc, "status": request.form["status"]})
            group_id = row("""SELECT id FROM journal_batch_groups
                              WHERE batch_id=:batch AND group_no=:group_no""",
                           {"batch": batch_id, "group_no": group_no})["id"]
            for line in lines:
                execute("""INSERT INTO journal_batch_lines(
                    group_id,account_id,debit,credit,taxable,tax_direction,party_type,
                    supplier_id,customer_id,tax_number,invoice_number,invoice_date,
                    line_description,cost_center_id)
                    VALUES(:group_id,:account_id,:debit,:credit,:taxable,:tax_direction,:party_type,
                    :supplier_id,:customer_id,:tax_number,:invoice_number,:invoice_date,
                    :line_description,:cost_center_id)""", {"group_id": group_id, **line})

        audit("CREATE", "JOURNAL_BATCH", f"إنشاء دفعة قيود {batch_no}")
        flash(f"تم حفظ دفعة القيود {batch_no}", "success")
        if request.form.get("print_after_save") == "1":
            return redirect(url_for("multi_journal_view", batch_id=batch_id, print=1))
        return redirect(url_for("multi_journal_view", batch_id=batch_id))

    batches = rows("SELECT * FROM journal_batches ORDER BY batch_date DESC,id DESC")
    return render_template(
        "multi_journal.html",
        batches=batches,
        accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                         WHERE active=1 AND accepts_entries=1 ORDER BY account_code"""),
        suppliers=rows("SELECT id,name,vat_number FROM suppliers ORDER BY name"),
        customers=rows("SELECT id,name,vat_number FROM customers ORDER BY name"),
        centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code")
    )

@app.route("/multi-journal/<int:batch_id>")
@login_required
def multi_journal_view(batch_id):
    batch = row("SELECT * FROM journal_batches WHERE id=:id", {"id": batch_id})
    if not batch:
        return "الدفعة غير موجودة", 404
    groups = rows("SELECT * FROM journal_batch_groups WHERE batch_id=:id ORDER BY group_no", {"id": batch_id})
    lines = rows("""
        SELECT l.*,g.group_no,g.journal_no,a.account_code,a.account_name_ar,
               s.name supplier_name,cus.name customer_name,
               cc.code cost_center_code,cc.name cost_center_name
        FROM journal_batch_lines l
        JOIN journal_batch_groups g ON g.id=l.group_id
        JOIN chart_of_accounts a ON a.id=l.account_id
        LEFT JOIN suppliers s ON s.id=l.supplier_id
        LEFT JOIN customers cus ON cus.id=l.customer_id
        LEFT JOIN cost_centers cc ON cc.id=l.cost_center_id
        WHERE g.batch_id=:id ORDER BY g.group_no,l.id
    """, {"id": batch_id})
    company = row("SELECT * FROM settings WHERE id=1")
    return render_template("multi_journal_view.html", batch=batch, groups=groups, lines=lines, company=company, auto_print=request.args.get("print")=="1")


@app.route("/multi-journal/<int:batch_id>/group/<int:group_no>/print")
@login_required
def multi_journal_group_print(batch_id, group_no):
    batch = row("SELECT * FROM journal_batches WHERE id=:id", {"id": batch_id})
    if not batch:
        return "الدفعة غير موجودة", 404

    group = row("""
        SELECT *
        FROM journal_batch_groups
        WHERE batch_id=:batch_id AND group_no=:group_no
    """, {"batch_id": batch_id, "group_no": group_no})
    if not group:
        return "القيد غير موجود داخل الدفعة", 404

    lines = rows("""
        SELECT l.*,a.account_code,a.account_name_ar,
               s.name supplier_name,cus.name customer_name,
               cc.code cost_center_code,cc.name cost_center_name
        FROM journal_batch_lines l
        JOIN chart_of_accounts a ON a.id=l.account_id
        LEFT JOIN suppliers s ON s.id=l.supplier_id
        LEFT JOIN customers cus ON cus.id=l.customer_id
        LEFT JOIN cost_centers cc ON cc.id=l.cost_center_id
        WHERE l.group_id=:group_id
        ORDER BY l.id
    """, {"group_id": group["id"]})

    company = row("SELECT * FROM settings WHERE id=1")
    return render_template(
        "multi_journal_group_print.html",
        batch=batch,
        group=group,
        lines=lines,
        company=company
    )


@app.route("/multi-journal/<int:batch_id>/delete", methods=["POST"])
@login_required
def multi_journal_delete(batch_id):
    batch = row("SELECT * FROM journal_batches WHERE id=:id", {"id": batch_id})
    if not batch:
        return "الدفعة غير موجودة", 404
    if batch["status"] == "مرحّل":
        flash("لا يمكن حذف دفعة مرحّلة", "danger")
    else:
        execute("DELETE FROM journal_batches WHERE id=:id", {"id": batch_id})
        audit("DELETE", "JOURNAL_BATCH", f"حذف دفعة {batch['batch_no']}")
        flash("تم حذف الدفعة", "success")
    return redirect(url_for("multi_journal"))

@app.route("/multi-journal/<int:batch_id>/export.xlsx")
@login_required
def multi_journal_export(batch_id):
    data = rows("""
        SELECT b.batch_no,b.batch_date,g.group_no,g.journal_no,a.account_code,a.account_name_ar,
               l.debit,l.credit,l.taxable,l.tax_direction,l.party_type,
               COALESCE(s.name,c.name) party_name,l.tax_number,l.invoice_number,l.invoice_date,
               l.line_description,cc.code cost_center_code,cc.name cost_center_name
        FROM journal_batches b
        JOIN journal_batch_groups g ON g.batch_id=b.id
        JOIN journal_batch_lines l ON l.group_id=g.id
        JOIN chart_of_accounts a ON a.id=l.account_id
        LEFT JOIN suppliers s ON s.id=l.supplier_id
        LEFT JOIN customers c ON c.id=l.customer_id
        LEFT JOIN cost_centers cc ON cc.id=l.cost_center_id
        WHERE b.id=:id ORDER BY g.group_no,l.id
    """, {"id": batch_id})
    headers = ["رقم الدفعة","تاريخ الدفعة","المجموعة","رقم القيد","الكود","اسم الحساب",
               "مدين","دائن","خاضع للضريبة","مدخلات/مخرجات","نوع الطرف","اسم الطرف",
               "الرقم الضريبي","رقم الفاتورة","تاريخ الفاتورة","البيان",
               "كود مركز التكلفة","مركز التكلفة"]
    keys = ["batch_no","batch_date","group_no","journal_no","account_code","account_name_ar",
            "debit","credit","taxable","tax_direction","party_type","party_name","tax_number",
            "invoice_number","invoice_date","line_description","cost_center_code","cost_center_name"]
    records = [[r.get(k) for k in keys] for r in data]
    return xlsx_response("multi_journal.xlsx", "قيود متعددة", headers, records)




ACCOUNTING_REPORTS = {
    "ledger-summary": {"no": 1, "title": "حساب الأستاذ - إجمالي", "group": "الحسابات"},
    "ledger-detail": {"no": 2, "title": "حساب الأستاذ - تفصيلي", "group": "الحسابات"},
    "cost-center-movements-center": {"no": 3, "title": "حركات مراكز التكلفة - حسب المركز", "group": "مراكز التكلفة"},
    "cost-center-movements-account": {"no": 4, "title": "حركات مراكز التكلفة - حسب الحساب", "group": "مراكز التكلفة"},
    "journal-report": {"no": 5, "title": "القيود المحاسبية", "group": "الحسابات"},
    "coa-report": {"no": 6, "title": "الدليل الحسابي", "group": "الحسابات"},
    "account-balances": {"no": 7, "title": "أرصدة الحسابات", "group": "الحسابات"},
    "trial-balance": {"no": 8, "title": "ميزان المراجعة", "group": "الحسابات"},
    "cc-ledger-summary": {"no": 9, "title": "حساب الأستاذ لمراكز التكلفة - إجمالي", "group": "مراكز التكلفة"},
    "cc-ledger-detail": {"no": 10, "title": "حساب الأستاذ لمراكز التكلفة - تفصيلي", "group": "مراكز التكلفة"},
    "cc-trial-balance": {"no": 11, "title": "ميزان المراجعة - مراكز التكلفة", "group": "مراكز التكلفة"},
    "cc-balances": {"no": 12, "title": "أرصدة مراكز التكلفة", "group": "مراكز التكلفة"},
    "accounts-with-centers": {"no": 13, "title": "إجمالي الحسابات مع مراكز التكلفة", "group": "تحليل مشترك"},
    "centers-with-accounts": {"no": 14, "title": "إجمالي مراكز التكلفة مع الحسابات", "group": "تحليل مشترك"},
    "vat-report": {"no": 15, "title": "ضريبة القيمة المضافة", "group": "الضرائب"},
}

def report_filters():
    return {
        "date_from": request.args.get("date_from", ""),
        "date_to": request.args.get("date_to", ""),
        "account_id": request.args.get("account_id", type=int),
        "cost_center_id": request.args.get("cost_center_id", type=int),
        "status": request.args.get("status", "مرحّل"),
        "tax_direction": request.args.get("tax_direction", ""),
        "q": request.args.get("q", "").strip(),
    }

def journal_filter_sql(filters, line_alias="l", journal_alias="j"):
    conditions = []
    params = {}
    if filters["date_from"]:
        conditions.append(f"{journal_alias}.journal_date >= :date_from")
        params["date_from"] = filters["date_from"]
    if filters["date_to"]:
        conditions.append(f"{journal_alias}.journal_date <= :date_to")
        params["date_to"] = filters["date_to"]
    if filters["status"] and filters["status"] != "الكل":
        conditions.append(f"{journal_alias}.status = :status")
        params["status"] = filters["status"]
    if filters["account_id"]:
        conditions.append(f"{line_alias}.account_id = :account_id")
        params["account_id"] = filters["account_id"]
    if filters["cost_center_id"]:
        conditions.append(f"{line_alias}.cost_center_id = :cost_center_id")
        params["cost_center_id"] = filters["cost_center_id"]
    if filters["tax_direction"]:
        conditions.append(f"{line_alias}.tax_direction = :tax_direction")
        params["tax_direction"] = filters["tax_direction"]
    if filters["q"]:
        conditions.append(f"""(
            {journal_alias}.journal_no ILIKE :q OR
            COALESCE({journal_alias}.reference,'') ILIKE :q OR
            COALESCE({journal_alias}.description,'') ILIKE :q OR
            COALESCE({line_alias}.invoice_number,'') ILIKE :q OR
            COALESCE({line_alias}.line_description,'') ILIKE :q
        )""")
        params["q"] = f"%{filters['q']}%"
    return (" WHERE " + " AND ".join(conditions)) if conditions else "", params

def report_context(slug, filters):
    definition = ACCOUNTING_REPORTS[slug]
    where, params = journal_filter_sql(filters)
    headers, data, summary = [], [], {}

    if slug == "coa-report":
        headers = ["الكود","اسم الحساب","الاسم الإنجليزي","النوع","الحساب الأب","المستوى","طبيعة الرصيد","القائمة المالية","يقبل حركة","الحالة"]
        data = rows("""
            SELECT a.id,a.account_code,a.account_name_ar,a.account_name_en,a.account_type,
                   p.account_name_ar parent_name,a.level,a.normal_balance,a.statement_type,
                   a.accepts_entries,a.active
            FROM chart_of_accounts a
            LEFT JOIN chart_of_accounts p ON p.id=a.parent_id
            ORDER BY a.account_code
        """)

    elif slug in ("ledger-summary","account-balances","trial-balance"):
        headers = ["الكود","اسم الحساب","إجمالي المدين","إجمالي الدائن","الرصيد المدين","الرصيد الدائن","عدد الحركات"]
        data = rows(f"""
            SELECT a.id,a.account_code,a.account_name_ar,
                   COALESCE(SUM(l.debit),0) total_debit,
                   COALESCE(SUM(l.credit),0) total_credit,
                   GREATEST(COALESCE(SUM(l.debit-l.credit),0),0) debit_balance,
                   GREATEST(COALESCE(SUM(l.credit-l.debit),0),0) credit_balance,
                   COUNT(l.id) movement_count
            FROM chart_of_accounts a
            LEFT JOIN journal_entry_lines l ON l.account_id=a.id
            LEFT JOIN journal_entries j ON j.id=l.journal_id
            {where}
            GROUP BY a.id,a.account_code,a.account_name_ar
            ORDER BY a.account_code
        """, params)
        summary = {
            "total_debit": sum(float(x["total_debit"]) for x in data),
            "total_credit": sum(float(x["total_credit"]) for x in data),
            "debit_balance": sum(float(x["debit_balance"]) for x in data),
            "credit_balance": sum(float(x["credit_balance"]) for x in data),
        }

    elif slug == "ledger-detail":
        headers = ["التاريخ","رقم القيد","المرجع","البيان","رقم الفاتورة","تاريخ الفاتورة","مدين","دائن","الرصيد المتحرك","الحالة"]
        raw = rows(f"""
            SELECT j.id journal_id,j.journal_date,j.journal_no,j.reference,j.description,
                   l.invoice_number,l.invoice_date,l.line_description,l.debit,l.credit,j.status
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            {where}
            ORDER BY j.journal_date,j.id,l.id
        """, params)
        running = 0.0
        data = []
        for item in raw:
            running += float(item["debit"]) - float(item["credit"])
            d = dict(item)
            d["running_balance"] = running
            data.append(d)
        summary = {
            "total_debit": sum(float(x["debit"]) for x in data),
            "total_credit": sum(float(x["credit"]) for x in data),
            "closing_balance": running,
        }

    elif slug == "journal-report":
        headers = ["تاريخ القيد","رقم القيد","الحالة","المرجع","البيان العام","الكود","اسم الحساب","مدين","دائن","اسم الطرف","الرقم الضريبي","رقم الفاتورة","تاريخ الفاتورة","مركز التكلفة"]
        data = rows(f"""
            SELECT j.id journal_id,j.journal_date,j.journal_no,j.status,j.reference,j.description,
                   a.account_code,a.account_name_ar,l.debit,l.credit,
                   COALESCE(s.name,cus.name,'') party_name,l.tax_number,l.invoice_number,l.invoice_date,
                   cc.code cost_center_code,cc.name cost_center_name
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            JOIN chart_of_accounts a ON a.id=l.account_id
            LEFT JOIN suppliers s ON s.id=l.supplier_id
            LEFT JOIN customers cus ON cus.id=l.customer_id
            LEFT JOIN cost_centers cc ON cc.id=l.cost_center_id
            {where}
            ORDER BY j.journal_date DESC,j.id DESC,l.id
        """, params)
        summary = {
            "total_debit": sum(float(x["debit"]) for x in data),
            "total_credit": sum(float(x["credit"]) for x in data),
            "count": len(data),
        }

    elif slug in ("cost-center-movements-center","cc-ledger-detail"):
        headers = ["مركز التكلفة","التاريخ","رقم القيد","الكود","اسم الحساب","البيان","مدين","دائن","الرصيد","الحالة"]
        raw = rows(f"""
            SELECT j.id journal_id,cc.code cost_center_code,cc.name cost_center_name,
                   j.journal_date,j.journal_no,j.status,a.account_code,a.account_name_ar,
                   COALESCE(l.line_description,j.description) line_description,l.debit,l.credit
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            JOIN chart_of_accounts a ON a.id=l.account_id
            JOIN cost_centers cc ON cc.id=l.cost_center_id
            {where}
            ORDER BY cc.code,j.journal_date,j.id,l.id
        """, params)
        balances = {}
        data = []
        for item in raw:
            key = item["cost_center_code"]
            balances[key] = balances.get(key, 0.0) + float(item["debit"]) - float(item["credit"])
            d = dict(item); d["running_balance"] = balances[key]; data.append(d)
        summary = {"total_debit": sum(float(x["debit"]) for x in data), "total_credit": sum(float(x["credit"]) for x in data)}

    elif slug == "cost-center-movements-account":
        headers = ["الكود","اسم الحساب","مركز التكلفة","تاريخ القيد","رقم القيد","البيان","مدين","دائن","الحالة"]
        data = rows(f"""
            SELECT j.id journal_id,a.account_code,a.account_name_ar,
                   cc.code cost_center_code,cc.name cost_center_name,
                   j.journal_date,j.journal_no,j.status,
                   COALESCE(l.line_description,j.description) line_description,l.debit,l.credit
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            JOIN chart_of_accounts a ON a.id=l.account_id
            JOIN cost_centers cc ON cc.id=l.cost_center_id
            {where}
            ORDER BY a.account_code,cc.code,j.journal_date,j.id
        """, params)

    elif slug in ("cc-ledger-summary","cc-trial-balance","cc-balances"):
        headers = ["كود المركز","مركز التكلفة","إجمالي المدين","إجمالي الدائن","الرصيد المدين","الرصيد الدائن","عدد الحركات"]
        data = rows(f"""
            SELECT cc.id,cc.code cost_center_code,cc.name cost_center_name,
                   COALESCE(SUM(l.debit),0) total_debit,
                   COALESCE(SUM(l.credit),0) total_credit,
                   GREATEST(COALESCE(SUM(l.debit-l.credit),0),0) debit_balance,
                   GREATEST(COALESCE(SUM(l.credit-l.debit),0),0) credit_balance,
                   COUNT(l.id) movement_count
            FROM cost_centers cc
            LEFT JOIN journal_entry_lines l ON l.cost_center_id=cc.id
            LEFT JOIN journal_entries j ON j.id=l.journal_id
            {where}
            GROUP BY cc.id,cc.code,cc.name
            ORDER BY cc.code
        """, params)
        summary = {
            "total_debit": sum(float(x["total_debit"]) for x in data),
            "total_credit": sum(float(x["total_credit"]) for x in data),
        }

    elif slug == "accounts-with-centers":
        headers = ["كود الحساب","اسم الحساب","كود المركز","مركز التكلفة","إجمالي المدين","إجمالي الدائن","الرصيد"]
        data = rows(f"""
            SELECT a.account_code,a.account_name_ar,cc.code cost_center_code,cc.name cost_center_name,
                   SUM(l.debit) total_debit,SUM(l.credit) total_credit,SUM(l.debit-l.credit) balance
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            JOIN chart_of_accounts a ON a.id=l.account_id
            LEFT JOIN cost_centers cc ON cc.id=l.cost_center_id
            {where}
            GROUP BY a.account_code,a.account_name_ar,cc.code,cc.name
            ORDER BY a.account_code,cc.code
        """, params)

    elif slug == "centers-with-accounts":
        headers = ["كود المركز","مركز التكلفة","كود الحساب","اسم الحساب","إجمالي المدين","إجمالي الدائن","الرصيد"]
        data = rows(f"""
            SELECT cc.code cost_center_code,cc.name cost_center_name,a.account_code,a.account_name_ar,
                   SUM(l.debit) total_debit,SUM(l.credit) total_credit,SUM(l.debit-l.credit) balance
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            JOIN chart_of_accounts a ON a.id=l.account_id
            LEFT JOIN cost_centers cc ON cc.id=l.cost_center_id
            {where}
            GROUP BY cc.code,cc.name,a.account_code,a.account_name_ar
            ORDER BY cc.code,a.account_code
        """, params)

    elif slug == "vat-report":
        vat_where, vat_params = journal_filter_sql(filters)
        vat_condition = "l.taxable=1"
        vat_where = vat_where + (" AND " if vat_where else " WHERE ") + vat_condition
        headers = ["تاريخ القيد","رقم القيد","تاريخ الفاتورة","رقم الفاتورة","اسم المورد / العميل","الرقم الضريبي","مدخلات / مخرجات","المبلغ قبل الضريبة","ضريبة القيمة المضافة","الإجمالي","حالة البيانات"]
        raw = rows(f"""
            SELECT j.id journal_id,j.journal_date,j.journal_no,l.invoice_date,l.invoice_number,
                   COALESCE(s.name,cus.name,'') party_name,l.tax_number,l.tax_direction,
                   GREATEST(l.debit,l.credit) amount_before_tax
            FROM journal_entry_lines l
            JOIN journal_entries j ON j.id=l.journal_id
            LEFT JOIN suppliers s ON s.id=l.supplier_id
            LEFT JOIN customers cus ON cus.id=l.customer_id
            {vat_where}
            ORDER BY COALESCE(l.invoice_date,j.journal_date),j.id,l.id
        """, vat_params)
        vat_rate = float(row("SELECT vat_rate FROM settings WHERE id=1")["vat_rate"] or 15)
        data = []
        for item in raw:
            d = dict(item)
            base_amount = float(item["amount_before_tax"] or 0)
            d["vat_amount"] = round(base_amount * vat_rate / 100, 2)
            d["total_amount"] = round(base_amount + d["vat_amount"], 2)
            complete = all([item["invoice_date"], item["invoice_number"], item["party_name"], item["tax_number"], item["tax_direction"] in ("مدخلات","مخرجات")])
            d["compliance_status"] = "مكتمل" if complete else "بيانات ناقصة"
            data.append(d)
        output_rows = [x for x in data if x["tax_direction"] == "مخرجات"]
        input_rows = [x for x in data if x["tax_direction"] == "مدخلات"]
        summary = {
            "output_base": sum(x["amount_before_tax"] for x in output_rows),
            "output_vat": sum(x["vat_amount"] for x in output_rows),
            "input_base": sum(x["amount_before_tax"] for x in input_rows),
            "input_vat": sum(x["vat_amount"] for x in input_rows),
            "net_vat": sum(x["vat_amount"] for x in output_rows) - sum(x["vat_amount"] for x in input_rows),
        }

    return {
        "definition": definition,
        "headers": headers,
        "data": data,
        "summary": summary,
    }





@app.route("/sales")
@login_required
def sales_center():
    return render_template("sales_center.html",
      quotations=row("SELECT COUNT(*) c FROM sales_quotations")["c"],
      orders=row("SELECT COUNT(*) c FROM sales_orders")["c"],
      deliveries=row("SELECT COUNT(*) c FROM sales_deliveries")["c"],
      invoices=row("SELECT COUNT(*) c FROM invoices")["c"])

@app.route("/sales/quotations",methods=["GET","POST"])
@login_required
def sales_quotations():
    if request.method=="POST":
        try: lines,sub,disc,vat,total=parse_sales_lines(request.form)
        except Exception as exc:
            flash(str(exc),"danger"); return redirect(url_for("sales_quotations"))
        no=next_sales_no("sales_quotations","quotation_date","QT",request.form["quotation_date"])
        execute("""INSERT INTO sales_quotations(quotation_no,quotation_date,valid_until,customer_id,
          branch_id,cost_center_id,status,subtotal,discount,vat,total,notes,created_by,created_at)
          VALUES(:no,:dt,:valid,:customer,:branch,:cc,:status,:sub,:disc,:vat,:total,:notes,:uid,:created)""",
          {"no":no,"dt":request.form["quotation_date"],"valid":request.form.get("valid_until") or None,
           "customer":request.form["customer_id"],"branch":request.form.get("branch_id") or None,
           "cc":request.form.get("cost_center_id") or None,"status":request.form.get("status","مسودة"),
           "sub":sub,"disc":disc,"vat":vat,"total":total,"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        qid=row("SELECT id FROM sales_quotations WHERE quotation_no=:n",{"n":no})["id"]
        for x in lines:
            execute("""INSERT INTO sales_quotation_items(quotation_id,item_id,item_name,quantity,unit,
              unit_price,discount_rate,vat_rate,line_subtotal,line_discount,line_vat,line_total)
              VALUES(:qid,:item_id,:item_name,:quantity,:unit,:unit_price,:discount_rate,:vat_rate,
              :line_subtotal,:line_discount,:line_vat,:line_total)""",{"qid":qid,**x})
        flash(f"تم إنشاء عرض السعر {no}","success")
        return redirect(url_for("sales_quotation_view",quotation_id=qid))
    return render_template("sales_quotations.html",
      docs=rows("""SELECT q.*,c.name customer_name FROM sales_quotations q JOIN customers c ON c.id=q.customer_id ORDER BY q.id DESC"""),
      customers=rows("SELECT id,name,name_en FROM customers ORDER BY name"),
      items=rows("SELECT id,sku,name,unit,sale_price,quantity FROM inventory WHERE active=1 ORDER BY name"),
      branches=rows("SELECT id,name FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"))

@app.route("/sales/quotations/<int:quotation_id>")
@login_required
def sales_quotation_view(quotation_id):
    doc=row("""SELECT q.*,c.name customer_name,c.name_en customer_name_en FROM sales_quotations q
               JOIN customers c ON c.id=q.customer_id WHERE q.id=:id""",{"id":quotation_id})
    items=rows("SELECT * FROM sales_quotation_items WHERE quotation_id=:id ORDER BY id",{"id":quotation_id})
    return render_template("sales_doc_view.html",doc=doc,items=items,title="عرض سعر / Quotation",
                           number=doc["quotation_no"],date=doc["quotation_date"],kind="quotation")

@app.route("/sales/quotations/<int:quotation_id>/convert",methods=["POST"])
@login_required
def sales_quotation_convert(quotation_id):
    q=row("SELECT * FROM sales_quotations WHERE id=:id",{"id":quotation_id})
    if q.get("converted_order_id"): return redirect(url_for("sales_order_view",order_id=q["converted_order_id"]))
    wh=row("SELECT id FROM warehouses WHERE active=1 ORDER BY id LIMIT 1")
    no=next_sales_no("sales_orders","order_date","SO",datetime.now().date().isoformat())
    execute("""INSERT INTO sales_orders(order_no,order_date,quotation_id,customer_id,branch_id,
      cost_center_id,warehouse_id,status,subtotal,discount,vat,total,notes,created_by,created_at)
      VALUES(:no,:dt,:qid,:customer,:branch,:cc,:wh,'مسودة',:sub,:disc,:vat,:total,:notes,:uid,:created)""",
      {"no":no,"dt":datetime.now().date(),"qid":quotation_id,"customer":q["customer_id"],
       "branch":q["branch_id"],"cc":q["cost_center_id"],"wh":wh["id"] if wh else None,
       "sub":q["subtotal"],"disc":q["discount"],"vat":q["vat"],"total":q["total"],
       "notes":f"من عرض السعر {q['quotation_no']}","uid":session.get("user_id"),"created":datetime.now()})
    oid=row("SELECT id FROM sales_orders WHERE order_no=:n",{"n":no})["id"]
    execute("""INSERT INTO sales_order_items(order_id,item_id,item_name,quantity,unit,unit_price,
      discount_rate,vat_rate,line_subtotal,line_discount,line_vat,line_total)
      SELECT :oid,item_id,item_name,quantity,unit,unit_price,discount_rate,vat_rate,
      line_subtotal,line_discount,line_vat,line_total FROM sales_quotation_items WHERE quotation_id=:qid""",
      {"oid":oid,"qid":quotation_id})
    execute("UPDATE sales_quotations SET status='محوّل',converted_order_id=:oid WHERE id=:id",
            {"oid":oid,"id":quotation_id})
    return redirect(url_for("sales_order_view",order_id=oid))

@app.route("/sales/orders",methods=["GET","POST"])
@login_required
def sales_orders():
    if request.method=="POST":
        try: lines,sub,disc,vat,total=parse_sales_lines(request.form)
        except Exception as exc:
            flash(str(exc),"danger"); return redirect(url_for("sales_orders"))
        no=next_sales_no("sales_orders","order_date","SO",request.form["order_date"])
        execute("""INSERT INTO sales_orders(order_no,order_date,customer_id,branch_id,cost_center_id,
          warehouse_id,status,subtotal,discount,vat,total,notes,created_by,created_at)
          VALUES(:no,:dt,:customer,:branch,:cc,:wh,:status,:sub,:disc,:vat,:total,:notes,:uid,:created)""",
          {"no":no,"dt":request.form["order_date"],"customer":request.form["customer_id"],
           "branch":request.form.get("branch_id") or None,"cc":request.form.get("cost_center_id") or None,
           "wh":request.form["warehouse_id"],"status":request.form.get("status","مسودة"),
           "sub":sub,"disc":disc,"vat":vat,"total":total,"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        oid=row("SELECT id FROM sales_orders WHERE order_no=:n",{"n":no})["id"]
        for x in lines:
            execute("""INSERT INTO sales_order_items(order_id,item_id,item_name,quantity,unit,unit_price,
              discount_rate,vat_rate,line_subtotal,line_discount,line_vat,line_total)
              VALUES(:oid,:item_id,:item_name,:quantity,:unit,:unit_price,:discount_rate,:vat_rate,
              :line_subtotal,:line_discount,:line_vat,:line_total)""",{"oid":oid,**x})
        return redirect(url_for("sales_order_view",order_id=oid))
    return render_template("sales_orders.html",
      docs=rows("""SELECT o.*,c.name customer_name,w.name warehouse_name FROM sales_orders o
                   JOIN customers c ON c.id=o.customer_id LEFT JOIN warehouses w ON w.id=o.warehouse_id ORDER BY o.id DESC"""),
      customers=rows("SELECT id,name,name_en FROM customers ORDER BY name"),
      items=rows("SELECT id,sku,name,unit,sale_price,quantity FROM inventory WHERE active=1 ORDER BY name"),
      branches=rows("SELECT id,name FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"),
      warehouses=rows("SELECT id,code,name FROM warehouses WHERE active=1 ORDER BY code"))

@app.route("/sales/orders/<int:order_id>")
@login_required
def sales_order_view(order_id):
    doc=row("""SELECT o.*,c.name customer_name,c.name_en customer_name_en,w.name warehouse_name
               FROM sales_orders o JOIN customers c ON c.id=o.customer_id
               LEFT JOIN warehouses w ON w.id=o.warehouse_id WHERE o.id=:id""",{"id":order_id})
    items=rows("SELECT * FROM sales_order_items WHERE order_id=:id ORDER BY id",{"id":order_id})
    return render_template("sales_doc_view.html",doc=doc,items=items,title="أمر بيع / Sales Order",
                           number=doc["order_no"],date=doc["order_date"],kind="order")

@app.route("/sales/orders/<int:order_id>/approve",methods=["POST"])
@login_required
def sales_order_approve(order_id):
    o=row("SELECT * FROM sales_orders WHERE id=:id",{"id":order_id})
    for i in rows("SELECT * FROM sales_order_items WHERE order_id=:id",{"id":order_id}):
        if warehouse_stock(i["item_id"],o["warehouse_id"]) < float(i["quantity"])-float(i["delivered_qty"]):
            flash(f"الرصيد غير كافٍ للصنف {i['item_name']}","danger")
            return redirect(url_for("sales_order_view",order_id=order_id))
    execute("UPDATE sales_orders SET status='معتمد' WHERE id=:id",{"id":order_id})
    return redirect(url_for("sales_order_view",order_id=order_id))

@app.route("/sales/deliveries",methods=["GET","POST"])
@login_required
def sales_deliveries():
    if request.method=="POST":
        oid=int(request.form["order_id"]); order=row("SELECT * FROM sales_orders WHERE id=:id",{"id":oid})
        item_ids=request.form.getlist("order_item_id[]"); qtys=request.form.getlist("delivery_qty[]")
        lines=[]
        for i,item_id in enumerate(item_ids):
            oi=row("SELECT * FROM sales_order_items WHERE id=:id",{"id":item_id})
            qty=float(qtys[i] or 0); remain=float(oi["quantity"])-float(oi["delivered_qty"])
            if qty>remain+0.0001: flash("كمية التسليم تتجاوز المتبقي","danger"); return redirect(url_for("sales_deliveries"))
            if qty>0:
                if warehouse_stock(oi["item_id"],order["warehouse_id"])<qty:
                    flash(f"الرصيد غير كافٍ للصنف {oi['item_name']}","danger"); return redirect(url_for("sales_deliveries"))
                lines.append((oi,qty))
        if not lines: flash("أدخل كمية تسليم","danger"); return redirect(url_for("sales_deliveries"))
        no=next_sales_no("sales_deliveries","delivery_date","DN",request.form["delivery_date"])
        execute("""INSERT INTO sales_deliveries(delivery_no,delivery_date,order_id,customer_id,
          warehouse_id,status,notes,created_by,created_at)
          VALUES(:no,:dt,:oid,:customer,:wh,'معتمد',:notes,:uid,:created)""",
          {"no":no,"dt":request.form["delivery_date"],"oid":oid,"customer":order["customer_id"],
           "wh":order["warehouse_id"],"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        did=row("SELECT id FROM sales_deliveries WHERE delivery_no=:n",{"n":no})["id"]
        for oi,qty in lines:
            item=row("SELECT cost FROM inventory WHERE id=:id",{"id":oi["item_id"]})
            execute("""INSERT INTO sales_delivery_items(delivery_id,order_item_id,item_id,quantity,unit_cost)
                       VALUES(:did,:oi,:item,:qty,:cost)""",
                    {"did":did,"oi":oi["id"],"item":oi["item_id"],"qty":qty,"cost":item["cost"]})
            execute("UPDATE sales_order_items SET delivered_qty=delivered_qty+:q WHERE id=:id",
                    {"q":qty,"id":oi["id"]})
            record_inventory_movement(request.form["delivery_date"],"بيع",oi["item_id"],order["warehouse_id"],
                                      qty,item["cost"],reference_type="DELIVERY",reference_id=did,reference_no=no)
        remaining=row("SELECT COUNT(*) c FROM sales_order_items WHERE order_id=:id AND delivered_qty<quantity",
                      {"id":oid})["c"]
        execute("UPDATE sales_orders SET status=:s WHERE id=:id",
                {"s":"مكتمل التسليم" if remaining==0 else "تسليم جزئي","id":oid})
        return redirect(url_for("sales_delivery_view",delivery_id=did))
    return render_template("sales_deliveries.html",
      orders=rows("""SELECT o.id,o.order_no,c.name customer_name FROM sales_orders o
                     JOIN customers c ON c.id=o.customer_id WHERE o.status IN ('معتمد','تسليم جزئي') ORDER BY o.id DESC"""),
      docs=rows("""SELECT d.*,o.order_no,c.name customer_name FROM sales_deliveries d
                   JOIN sales_orders o ON o.id=d.order_id JOIN customers c ON c.id=d.customer_id ORDER BY d.id DESC"""))

@app.route("/sales/orders/<int:order_id>/items")
@login_required
def sales_order_items_api(order_id):
    return {"items":[dict(x) for x in rows("""SELECT id,item_name,quantity,delivered_qty
      FROM sales_order_items WHERE order_id=:id ORDER BY id""",{"id":order_id})]}

@app.route("/sales/deliveries/<int:delivery_id>")
@login_required
def sales_delivery_view(delivery_id):
    doc=row("""SELECT d.*,o.order_no,c.name customer_name,c.name_en customer_name_en,w.name warehouse_name
               FROM sales_deliveries d JOIN sales_orders o ON o.id=d.order_id
               JOIN customers c ON c.id=d.customer_id JOIN warehouses w ON w.id=d.warehouse_id
               WHERE d.id=:id""",{"id":delivery_id})
    items=rows("""SELECT di.*,oi.item_name,oi.unit FROM sales_delivery_items di
                  JOIN sales_order_items oi ON oi.id=di.order_item_id WHERE di.delivery_id=:id""",{"id":delivery_id})
    return render_template("sales_delivery_view.html",doc=doc,items=items)


@app.route("/sales/deliveries/<int:delivery_id>/create-invoice",methods=["POST"])
@login_required
def sales_delivery_create_invoice(delivery_id):
    try:
        invoice_id = create_invoice_from_sales_delivery(delivery_id)
        post_invoice_to_ledger(invoice_id)
        execute("UPDATE invoices SET status='معتمدة' WHERE id=:id",{"id":invoice_id})
        flash("تم إنشاء الفاتورة وترحيل قيد المبيعات وتكلفة البضاعة المباعة","success")
        return redirect(url_for("invoice_view",invoice_id=invoice_id))
    except Exception as exc:
        db.session.rollback()
        flash(str(exc),"danger")
        return redirect(url_for("sales_delivery_view",delivery_id=delivery_id))

@app.route("/sales/returns",methods=["GET","POST"])
@login_required
def sales_returns():
    if request.method=="POST":
        invoice_id=int(request.form["invoice_id"])
        invoice=row("SELECT * FROM invoices WHERE id=:id",{"id":invoice_id})
        if not invoice:
            flash("الفاتورة غير موجودة","danger")
            return redirect(url_for("sales_returns"))
        warehouse_id=int(request.form["warehouse_id"])
        ids=request.form.getlist("invoice_item_id[]")
        quantities=request.form.getlist("return_qty[]")
        prepared=[]; subtotal=vat_total=grand_total=cost_total=0.0
        for index,item_id in enumerate(ids):
            item=row("""SELECT ii.*,COALESCE(inv.cost,0) current_cost
                        FROM invoice_items ii LEFT JOIN inventory inv ON inv.id=ii.item_id
                        WHERE ii.id=:id AND ii.invoice_id=:invoice""",
                     {"id":item_id,"invoice":invoice_id})
            if not item:
                continue
            qty=float(quantities[index] or 0)
            previous=db.session.execute(text("""SELECT COALESCE(SUM(sri.quantity),0)
                FROM sales_return_items sri JOIN sales_returns sr ON sr.id=sri.return_id
                WHERE sri.invoice_item_id=:item AND sr.invoice_id=:invoice"""),
                {"item":item_id,"invoice":invoice_id}).scalar() or 0
            available=float(item["quantity"])-float(previous)
            if qty < 0 or qty > available + 0.0001:
                flash(f"كمية المرتجع للصنف {item['item_name']} تتجاوز المتاح {available:.3f}","danger")
                return redirect(url_for("sales_returns"))
            if qty == 0:
                continue
            base=round(qty*float(item["unit_price"]),2)
            tax=round(base*float(item["vat_rate"] or 0)/100,2)
            cost=round(qty*float(item["current_cost"] or 0),2)
            prepared.append((item,qty,base,tax,cost))
            subtotal+=base;vat_total+=tax;grand_total+=base+tax;cost_total+=cost
        if not prepared:
            flash("أدخل كمية مرتجع واحدة على الأقل","danger")
            return redirect(url_for("sales_returns"))

        no=next_sales_no("sales_returns","return_date","SR",request.form["return_date"]) \
           if "sales_returns" in {"sales_returns"} else ""
        # next_sales_no does not originally include returns; generate safely.
        year=datetime.strptime(request.form["return_date"],"%Y-%m-%d").year
        seq=db.session.execute(text("SELECT COUNT(*) FROM sales_returns WHERE EXTRACT(YEAR FROM return_date)=:y"),
                               {"y":year}).scalar() or 0
        no=f"SR-{year}-{seq+1:06d}"

        execute("""INSERT INTO sales_returns(
          return_no,return_date,invoice_id,customer_id,warehouse_id,subtotal,vat,total,
          cost_total,status,notes,created_by,created_at)
          VALUES(:no,:dt,:invoice,:customer,:warehouse,:sub,:vat,:total,:cost,
          'معتمد',:notes,:uid,:created)""",
          {"no":no,"dt":request.form["return_date"],"invoice":invoice_id,
           "customer":invoice["customer_id"],"warehouse":warehouse_id,
           "sub":round(subtotal,2),"vat":round(vat_total,2),"total":round(grand_total,2),
           "cost":round(cost_total,2),"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        return_id=row("SELECT id FROM sales_returns WHERE return_no=:no",{"no":no})["id"]

        for item,qty,base,tax,cost in prepared:
            unit_cost=float(item["current_cost"] or 0)
            execute("""INSERT INTO sales_return_items(
              return_id,invoice_item_id,item_id,item_name,quantity,unit,unit_price,vat_rate,
              unit_cost,line_subtotal,line_vat,line_total)
              VALUES(:return_id,:invoice_item,:item_id,:name,:qty,:unit,:price,:rate,
              :cost,:sub,:vat,:total)""",
              {"return_id":return_id,"invoice_item":item["id"],"item_id":item["item_id"],
               "name":item["item_name"],"qty":qty,"unit":item["unit"],
               "price":item["unit_price"],"rate":item["vat_rate"],"cost":unit_cost,
               "sub":base,"vat":tax,"total":round(base+tax,2)})
            if item["item_id"]:
                record_inventory_movement(
                    request.form["return_date"],"مرتجع مبيعات",item["item_id"],warehouse_id,
                    qty,unit_cost,reference_type="SALES_RETURN",reference_id=return_id,
                    reference_no=no,notes=f"مرتجع للفاتورة {invoice['invoice_no']}"
                )
        post_sales_return_to_ledger(return_id)
        audit("CREATE","SALES_RETURN",f"إنشاء مرتجع مبيعات {no}")
        flash(f"تم إنشاء المرتجع {no} وإعادة المخزون وترحيل القيد","success")
        return redirect(url_for("sales_return_view",return_id=return_id))

    return render_template("sales_returns.html",
      invoices=rows("""SELECT i.id,i.invoice_no,i.invoice_date,c.name customer_name
                       FROM invoices i JOIN customers c ON c.id=i.customer_id
                       WHERE i.status='معتمدة' ORDER BY i.id DESC"""),
      warehouses=rows("SELECT id,code,name FROM warehouses WHERE active=1 ORDER BY code"),
      docs=rows("""SELECT r.*,i.invoice_no,c.name customer_name,j.journal_no
                   FROM sales_returns r JOIN invoices i ON i.id=r.invoice_id
                   JOIN customers c ON c.id=r.customer_id
                   LEFT JOIN journal_entries j ON j.id=r.journal_id ORDER BY r.id DESC"""))

@app.route("/sales/invoices/<int:invoice_id>/return-items")
@login_required
def sales_invoice_return_items(invoice_id):
    data=rows("""SELECT ii.id,ii.item_id,ii.item_name,ii.quantity,ii.unit,ii.unit_price,ii.vat_rate,
      ii.quantity-COALESCE((SELECT SUM(sri.quantity) FROM sales_return_items sri
      JOIN sales_returns sr ON sr.id=sri.return_id
      WHERE sri.invoice_item_id=ii.id AND sr.invoice_id=:invoice),0) returnable_qty
      FROM invoice_items ii WHERE ii.invoice_id=:invoice ORDER BY ii.id""",{"invoice":invoice_id})
    return {"items":[dict(x) for x in data]}

@app.route("/sales/returns/<int:return_id>")
@login_required
def sales_return_view(return_id):
    doc=row("""SELECT r.*,i.invoice_no,c.name customer_name,c.name_en customer_name_en,
                      w.name warehouse_name,j.journal_no
               FROM sales_returns r JOIN invoices i ON i.id=r.invoice_id
               JOIN customers c ON c.id=r.customer_id
               JOIN warehouses w ON w.id=r.warehouse_id
               LEFT JOIN journal_entries j ON j.id=r.journal_id
               WHERE r.id=:id""",{"id":return_id})
    if not doc:
        return "مرتجع المبيعات غير موجود",404
    items=rows("SELECT * FROM sales_return_items WHERE return_id=:id ORDER BY id",{"id":return_id})
    return render_template("sales_return_view.html",doc=doc,items=items)

@app.route("/sales/dashboard")
@login_required
def sales_dashboard():
    summary=row("""SELECT
      COALESCE(SUM(total),0) total_sales,
      COALESCE(SUM(vat),0) total_vat,
      COUNT(*) invoice_count
      FROM invoices WHERE status='معتمدة'""")
    return_summary=row("""SELECT COALESCE(SUM(total),0) total_returns,
                                  COUNT(*) return_count FROM sales_returns""")
    profit=row("""SELECT COALESCE(SUM(i.total),0)-COALESCE((SELECT SUM(r.total) FROM sales_returns r),0) net_sales,
      COALESCE((SELECT SUM(di.quantity*di.unit_cost) FROM sales_delivery_items di),0)
      -COALESCE((SELECT SUM(COALESCE(sri.quantity,0)*COALESCE(sri.unit_cost,0))
                 FROM sales_return_items sri),0) net_cost""")
    monthly=rows("""SELECT TO_CHAR(invoice_date,'YYYY-MM') month,
                    SUM(total) sales,COUNT(*) invoice_count
                    FROM invoices WHERE status='معتمدة'
                    GROUP BY TO_CHAR(invoice_date,'YYYY-MM') ORDER BY month DESC LIMIT 12""")
    customers=rows("""SELECT c.name,SUM(i.total) total
                      FROM invoices i JOIN customers c ON c.id=i.customer_id
                      WHERE i.status='معتمدة' GROUP BY c.id,c.name
                      ORDER BY total DESC LIMIT 10""")
    items=rows("""SELECT ii.item_name,SUM(ii.quantity) quantity,SUM(ii.line_total) sales
                  FROM invoice_items ii JOIN invoices i ON i.id=ii.invoice_id
                  WHERE i.status='معتمدة' GROUP BY ii.item_name
                  ORDER BY sales DESC LIMIT 10""")
    return render_template("sales_dashboard.html",summary=summary,
      return_summary=return_summary,profit=profit,monthly=monthly,
      top_customers=customers,top_items=items)

@app.route("/sales/reports")
@login_required
def sales_reports():
    date_from=request.args.get("date_from","")
    date_to=request.args.get("date_to","")
    customer_id=request.args.get("customer_id","")
    conditions=["i.status='معتمدة'"];params={}
    if date_from:
        conditions.append("i.invoice_date>=:date_from");params["date_from"]=date_from
    if date_to:
        conditions.append("i.invoice_date<=:date_to");params["date_to"]=date_to
    if customer_id:
        conditions.append("i.customer_id=:customer");params["customer"]=customer_id
    where=" AND ".join(conditions)
    data=rows(f"""SELECT i.*,c.name customer_name,
      COALESCE((SELECT SUM(di.quantity*di.unit_cost) FROM sales_delivery_items di
      JOIN sales_deliveries d ON d.id=di.delivery_id WHERE d.id=i.delivery_id),0) cost,
      i.subtotal-COALESCE((SELECT SUM(di.quantity*di.unit_cost) FROM sales_delivery_items di
      JOIN sales_deliveries d ON d.id=di.delivery_id WHERE d.id=i.delivery_id),0) gross_profit
      FROM invoices i JOIN customers c ON c.id=i.customer_id
      WHERE {where} ORDER BY i.invoice_date DESC,i.id DESC""",params)
    return render_template("sales_reports.html",rows=data,
      customers=rows("SELECT id,name FROM customers ORDER BY name"),
      date_from=date_from,date_to=date_to,customer_id=customer_id)

@app.route("/sales/reports.xlsx")
@login_required
def sales_reports_export():
    data=rows("""SELECT i.invoice_no,i.invoice_date,c.name customer_name,i.subtotal,i.vat,i.total,
      COALESCE((SELECT SUM(di.quantity*di.unit_cost) FROM sales_delivery_items di
      JOIN sales_deliveries d ON d.id=di.delivery_id WHERE d.id=i.delivery_id),0) cost
      FROM invoices i JOIN customers c ON c.id=i.customer_id
      WHERE i.status='معتمدة' ORDER BY i.invoice_date DESC,i.id DESC""")
    return xlsx_response("sales_profit_report.xlsx","المبيعات والأرباح",
      ["رقم الفاتورة","التاريخ","العميل","قبل الضريبة","الضريبة","الإجمالي","التكلفة","مجمل الربح"],
      [[r["invoice_no"],r["invoice_date"],r["customer_name"],r["subtotal"],r["vat"],
        r["total"],r["cost"],float(r["subtotal"])-float(r["cost"])] for r in data])


@app.route("/procurement")
@login_required
def procurement_center():
    stats = {
        "requisitions": row("SELECT COUNT(*) count FROM purchase_requisitions")["count"],
        "orders": row("SELECT COUNT(*) count FROM purchase_orders")["count"],
        "receipts": row("SELECT COUNT(*) count FROM goods_receipts")["count"],
        "supplier_invoices": row("SELECT COUNT(*) count FROM supplier_invoices")["count"],
    }
    return render_template("procurement_center.html",stats=stats)

@app.route("/purchase-requisitions",methods=["GET","POST"])
@login_required
def purchase_requisitions():
    if request.method=="POST":
        item_names=request.form.getlist("item_name[]")
        quantities=request.form.getlist("quantity[]")
        units=request.form.getlist("unit[]")
        prices=request.form.getlist("estimated_price[]")
        descriptions=request.form.getlist("description[]")
        required_dates=request.form.getlist("required_date[]")
        supplier_ids=request.form.getlist("suggested_supplier_id[]")
        item_codes=request.form.getlist("item_code[]")
        items=[]; total=0
        for i,name in enumerate(item_names):
            if not name.strip(): continue
            qty=float(quantities[i] or 0); price=float(prices[i] or 0)
            if qty<=0: continue
            total += qty*price
            items.append({"name":name.strip(),"qty":qty,"unit":units[i],
                          "price":price,"description":descriptions[i],
                          "required_date":required_dates[i] or None,
                          "supplier_id":supplier_ids[i] or None,
                          "item_code":item_codes[i]})
        if not items:
            flash("أدخل صنفًا واحدًا على الأقل","danger")
            return redirect(url_for("purchase_requisitions"))
        no=next_document_number("purchase_requisitions","requisition_date","PR",request.form["requisition_date"])
        execute("""INSERT INTO purchase_requisitions(
          requisition_no,requisition_date,branch_id,cost_center_id,requested_by,
          department,priority,reason,notes,status,total_estimated,created_by,created_at)
          VALUES(:no,:dt,:branch,:cc,:requested,:department,:priority,:reason,
          :notes,:status,:total,:uid,:created)""",
          {"no":no,"dt":request.form["requisition_date"],
           "branch":request.form.get("branch_id") or None,
           "cc":request.form.get("cost_center_id") or None,
           "requested":request.form.get("requested_by",""),
           "department":request.form.get("department",""),
           "priority":request.form.get("priority","عادية"),
           "reason":request.form.get("reason",""),
           "notes":request.form.get("notes",""),
           "status":request.form.get("status","مسودة"),
           "total":round(total,2),"uid":session.get("user_id"),"created":datetime.now()})
        req_id=row("SELECT id FROM purchase_requisitions WHERE requisition_no=:n",{"n":no})["id"]
        for x in items:
            execute("""INSERT INTO purchase_requisition_items(
              requisition_id,item_code,item_name,description,quantity,unit,
              estimated_price,required_date,suggested_supplier_id)
              VALUES(:rid,:code,:name,:des,:qty,:unit,:price,:rd,:sid)""",
              {"rid":req_id,"code":x["item_code"],"name":x["name"],"des":x["description"],
               "qty":x["qty"],"unit":x["unit"],"price":x["price"],
               "rd":x["required_date"],"sid":x["supplier_id"]})
        audit("CREATE","PURCHASE_REQUISITION",f"إنشاء طلب شراء {no}")
        flash(f"تم إنشاء طلب الشراء {no}","success")
        return redirect(url_for("purchase_requisition_view",req_id=req_id))
    reqs=rows("""SELECT pr.*,b.name branch_name,cc.name cost_center_name
                 FROM purchase_requisitions pr
                 LEFT JOIN branches b ON b.id=pr.branch_id
                 LEFT JOIN cost_centers cc ON cc.id=pr.cost_center_id
                 ORDER BY pr.requisition_date DESC,pr.id DESC""")
    return render_template("purchase_requisitions.html",reqs=reqs,
      branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT * FROM cost_centers WHERE active=1 ORDER BY code"),
      suppliers=rows("SELECT id,name,name_en FROM suppliers ORDER BY name"))

@app.route("/purchase-requisitions/<int:req_id>")
@login_required
def purchase_requisition_view(req_id):
    req=row("""SELECT pr.*,b.name branch_name,cc.name cost_center_name
               FROM purchase_requisitions pr
               LEFT JOIN branches b ON b.id=pr.branch_id
               LEFT JOIN cost_centers cc ON cc.id=pr.cost_center_id
               WHERE pr.id=:id""",{"id":req_id})
    if not req:return "طلب الشراء غير موجود",404
    items=rows("""SELECT i.*,s.name supplier_name FROM purchase_requisition_items i
                  LEFT JOIN suppliers s ON s.id=i.suggested_supplier_id
                  WHERE i.requisition_id=:id ORDER BY i.id""",{"id":req_id})
    return render_template("purchase_requisition_view.html",req=req,items=items,
                           approver=get_required_approver("طلب شراء",float(req["total_estimated"])))

@app.route("/purchase-requisitions/<int:req_id>/approve",methods=["POST"])
@login_required
def purchase_requisition_approve(req_id):
    execute("""UPDATE purchase_requisitions SET status='معتمد',
               approved_by=:u,approved_at=:t WHERE id=:id""",
            {"u":session.get("user_id"),"t":datetime.now(),"id":req_id})
    audit("APPROVE","PURCHASE_REQUISITION",f"اعتماد طلب شراء {req_id}")
    flash("تم اعتماد طلب الشراء","success")
    return redirect(url_for("purchase_requisition_view",req_id=req_id))

@app.route("/purchase-orders",methods=["GET","POST"])
@login_required
def purchase_orders():
    if request.method=="POST":
        req_id=request.form.get("requisition_id") or None
        if req_id:
            req=row("SELECT status FROM purchase_requisitions WHERE id=:id",{"id":req_id})
            if not req or req["status"]!="معتمد":
                flash("لا يمكن إصدار أمر شراء من طلب غير معتمد","danger")
                return redirect(url_for("purchase_orders"))
        names=request.form.getlist("item_name[]")
        qtys=request.form.getlist("quantity[]")
        prices=request.form.getlist("unit_price[]")
        vats=request.form.getlist("vat_rate[]")
        units=request.form.getlist("unit[]")
        codes=request.form.getlist("item_code[]")
        descriptions=request.form.getlist("description[]")
        items=[]; subtotal=0; vat=0
        for i,n in enumerate(names):
            if not n.strip(): continue
            q=float(qtys[i] or 0); p=float(prices[i] or 0); vr=float(vats[i] or 0)
            if q<=0:continue
            base=q*p; tax=base*vr/100
            subtotal+=base; vat+=tax
            items.append({"name":n.strip(),"qty":q,"price":p,"vat":vr,
                          "unit":units[i],"code":codes[i],"description":descriptions[i]})
        if not items:
            flash("أدخل صنفًا واحدًا على الأقل","danger")
            return redirect(url_for("purchase_orders"))
        no=next_document_number("purchase_orders","po_date","PO",request.form["po_date"])
        execute("""INSERT INTO purchase_orders(
          po_no,po_date,requisition_id,supplier_id,branch_id,cost_center_id,
          payment_terms,delivery_terms,warehouse,notes,status,subtotal,vat,total,
          created_by,created_at)
          VALUES(:no,:dt,:req,:supplier,:branch,:cc,:pay,:delivery,:warehouse,
          :notes,:status,:sub,:vat,:total,:uid,:created)""",
          {"no":no,"dt":request.form["po_date"],"req":req_id,
           "supplier":request.form["supplier_id"],"branch":request.form.get("branch_id") or None,
           "cc":request.form.get("cost_center_id") or None,
           "pay":request.form.get("payment_terms",""),
           "delivery":request.form.get("delivery_terms",""),
           "warehouse":request.form.get("warehouse",""),
           "notes":request.form.get("notes",""),
           "status":request.form.get("status","مسودة"),
           "sub":round(subtotal,2),"vat":round(vat,2),"total":round(subtotal+vat,2),
           "uid":session.get("user_id"),"created":datetime.now()})
        po_id=row("SELECT id FROM purchase_orders WHERE po_no=:n",{"n":no})["id"]
        for x in items:
            execute("""INSERT INTO purchase_order_items(
              po_id,item_code,item_name,description,quantity,unit,unit_price,vat_rate)
              VALUES(:po,:code,:name,:des,:qty,:unit,:price,:vat)""",
              {"po":po_id,"code":x["code"],"name":x["name"],"des":x["description"],
               "qty":x["qty"],"unit":x["unit"],"price":x["price"],"vat":x["vat"]})
        if req_id:
            execute("UPDATE purchase_requisitions SET status='تم إصدار أمر شراء' WHERE id=:id",{"id":req_id})
        audit("CREATE","PURCHASE_ORDER",f"إنشاء أمر شراء {no}")
        flash(f"تم إنشاء أمر الشراء {no}","success")
        return redirect(url_for("purchase_order_view",po_id=po_id))
    return render_template("purchase_orders.html",
      orders=rows("""SELECT po.*,s.name supplier_name FROM purchase_orders po
                     JOIN suppliers s ON s.id=po.supplier_id
                     ORDER BY po.po_date DESC,po.id DESC"""),
      requisitions=rows("""SELECT id,requisition_no,total_estimated FROM purchase_requisitions
                           WHERE status='معتمد' ORDER BY requisition_date DESC"""),
      suppliers=rows("SELECT id,name,name_en,vat_number FROM suppliers ORDER BY name"),
      branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT * FROM cost_centers WHERE active=1 ORDER BY code"))

@app.route("/purchase-orders/<int:po_id>")
@login_required
def purchase_order_view(po_id):
    po=row("""SELECT po.*,s.name supplier_name,s.name_en supplier_name_en,
              s.vat_number supplier_vat,b.name branch_name,cc.name cost_center_name
              FROM purchase_orders po JOIN suppliers s ON s.id=po.supplier_id
              LEFT JOIN branches b ON b.id=po.branch_id
              LEFT JOIN cost_centers cc ON cc.id=po.cost_center_id WHERE po.id=:id""",{"id":po_id})
    if not po:return "أمر الشراء غير موجود",404
    items=rows("SELECT * FROM purchase_order_items WHERE po_id=:id ORDER BY id",{"id":po_id})
    company=row("SELECT * FROM settings WHERE id=1")
    return render_template("purchase_order_view.html",po=po,items=items,company=company)

@app.route("/purchase-orders/<int:po_id>/approve",methods=["POST"])
@login_required
def purchase_order_approve(po_id):
    execute("""UPDATE purchase_orders SET status='معتمد',
               approved_by=:u,approved_at=:t WHERE id=:id""",
            {"u":session.get("user_id"),"t":datetime.now(),"id":po_id})
    flash("تم اعتماد أمر الشراء","success")
    return redirect(url_for("purchase_order_view",po_id=po_id))

@app.route("/goods-receipts",methods=["GET","POST"])
@login_required
def goods_receipts():
    if request.method=="POST":
        po_id=int(request.form["po_id"])
        po=row("SELECT * FROM purchase_orders WHERE id=:id",{"id":po_id})
        if not po or po["status"]!="معتمد":
            flash("يجب اعتماد أمر الشراء أولًا","danger")
            return redirect(url_for("goods_receipts"))
        item_ids=request.form.getlist("po_item_id[]")
        qtys=request.form.getlist("received_qty[]")
        accepted=request.form.getlist("accepted_qty[]")
        rejected=request.form.getlist("rejected_qty[]")
        notes=request.form.getlist("line_notes[]")
        lines=[]
        for i,item_id in enumerate(item_ids):
            item=row("SELECT * FROM purchase_order_items WHERE id=:id",{"id":item_id})
            if not item:continue
            rq=float(qtys[i] or 0); aq=float(accepted[i] or 0); rej=float(rejected[i] or 0)
            remaining=float(item["quantity"])-float(item["received_qty"])
            if rq<0 or rq>remaining+1e-9:
                flash(f"الكمية المستلمة للصنف {item['item_name']} تتجاوز المتبقي","danger")
                return redirect(url_for("goods_receipts"))
            if abs((aq+rej)-rq)>0.001:
                flash("المقبول + المرفوض يجب أن يساوي المستلم","danger")
                return redirect(url_for("goods_receipts"))
            if rq>0:lines.append((item_id,rq,aq,rej,notes[i]))
        if not lines:
            flash("أدخل كمية استلام","danger")
            return redirect(url_for("goods_receipts"))
        no=next_document_number("goods_receipts","grn_date","GRN",request.form["grn_date"])
        execute("""INSERT INTO goods_receipts(
          grn_no,grn_date,po_id,supplier_id,warehouse,notes,status,created_by,created_at)
          VALUES(:no,:dt,:po,:supplier,:wh,:notes,'معتمد',:uid,:created)""",
          {"no":no,"dt":request.form["grn_date"],"po":po_id,"supplier":po["supplier_id"],
           "wh":request.form.get("warehouse",""),"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        grn_id=row("SELECT id FROM goods_receipts WHERE grn_no=:n",{"n":no})["id"]
        for item_id,rq,aq,rej,note in lines:
            execute("""INSERT INTO goods_receipt_items(
              grn_id,po_item_id,received_qty,accepted_qty,rejected_qty,notes)
              VALUES(:grn,:item,:r,:a,:rej,:notes)""",
              {"grn":grn_id,"item":item_id,"r":rq,"a":aq,"rej":rej,"notes":note})
            execute("""UPDATE purchase_order_items SET received_qty=received_qty+:q
                       WHERE id=:id""",{"q":rq,"id":item_id})
            if aq>0:
                inv_item=row("""SELECT id,cost FROM inventory
                                WHERE (NULLIF(:sku,'') IS NOT NULL AND sku=:sku)
                                   OR LOWER(name)=LOWER(:name)
                                ORDER BY CASE WHEN sku=:sku THEN 0 ELSE 1 END LIMIT 1""",
                             {"sku":item["item_code"] or "","name":item["item_name"]})
                if not inv_item:
                    execute("""INSERT INTO inventory(sku,name,quantity,unit,cost,sale_price,reorder_level,active)
                               VALUES(:sku,:name,0,:unit,:cost,0,0,1)""",
                            {"sku":item["item_code"] or None,"name":item["item_name"],
                             "unit":item["unit"],"cost":item["unit_price"]})
                    inv_item=row("SELECT id,cost FROM inventory WHERE name=:name ORDER BY id DESC LIMIT 1",
                                 {"name":item["item_name"]})
                warehouse=row("SELECT id FROM warehouses WHERE active=1 ORDER BY id LIMIT 1")
                if warehouse:
                    record_inventory_movement(
                      request.form["grn_date"],"استلام",inv_item["id"],warehouse["id"],aq,
                      item["unit_price"],reference_type="GRN",reference_id=grn_id,
                      reference_no=no,notes=f"استلام أمر الشراء {po['po_no']}")

        remaining=row("""SELECT COUNT(*) count FROM purchase_order_items
                         WHERE po_id=:po AND received_qty<quantity""",{"po":po_id})["count"]
        execute("UPDATE purchase_orders SET status=:s WHERE id=:id",
                {"s":"مستلم بالكامل" if remaining==0 else "مستلم جزئيًا","id":po_id})
        audit("CREATE","GOODS_RECEIPT",f"إنشاء استلام {no}")
        flash(f"تم إنشاء استلام المواد {no}","success")
        return redirect(url_for("goods_receipt_view",grn_id=grn_id))
    return render_template("goods_receipts.html",
      pos=rows("""SELECT po.id,po.po_no,po.po_date,s.name supplier_name
                  FROM purchase_orders po JOIN suppliers s ON s.id=po.supplier_id
                  WHERE po.status IN ('معتمد','مستلم جزئيًا') ORDER BY po.po_date DESC"""),
      receipts=rows("""SELECT g.*,po.po_no,s.name supplier_name FROM goods_receipts g
                       JOIN purchase_orders po ON po.id=g.po_id
                       JOIN suppliers s ON s.id=g.supplier_id
                       ORDER BY g.grn_date DESC,g.id DESC"""))

@app.route("/purchase-orders/<int:po_id>/open-items")
@login_required
def purchase_order_open_items(po_id):
    po=row("SELECT * FROM purchase_orders WHERE id=:id",{"id":po_id})
    items=rows("""SELECT id,item_code,item_name,quantity,received_qty,unit
                  FROM purchase_order_items WHERE po_id=:id ORDER BY id""",{"id":po_id})
    return {"po":dict(po) if po else None,"items":[dict(x) for x in items]}

@app.route("/goods-receipts/<int:grn_id>")
@login_required
def goods_receipt_view(grn_id):
    grn=row("""SELECT g.*,po.po_no,s.name supplier_name FROM goods_receipts g
               JOIN purchase_orders po ON po.id=g.po_id
               JOIN suppliers s ON s.id=g.supplier_id WHERE g.id=:id""",{"id":grn_id})
    if not grn:return "الاستلام غير موجود",404
    items=rows("""SELECT gi.*,pi.item_code,pi.item_name,pi.unit
                  FROM goods_receipt_items gi JOIN purchase_order_items pi ON pi.id=gi.po_item_id
                  WHERE gi.grn_id=:id ORDER BY gi.id""",{"id":grn_id})
    return render_template("goods_receipt_view.html",grn=grn,items=items)

@app.route("/supplier-invoices",methods=["GET","POST"])
@login_required
def supplier_invoices():
    if request.method=="POST":
        subtotal=round(float(request.form["subtotal"] or 0),2)
        vat=round(float(request.form["vat"] or 0),2)
        total=round(subtotal+vat,2)
        internal=next_document_number("supplier_invoices","invoice_date","SI",request.form["invoice_date"])
        try:
            execute("""INSERT INTO supplier_invoices(
              supplier_invoice_no,internal_no,invoice_date,supplier_id,po_id,grn_id,
              cost_center_id,expense_or_inventory_account_id,subtotal,vat,total,status,
              posting_status,notes,created_by,created_at)
              VALUES(:supplier_no,:internal,:dt,:supplier,:po,:grn,:cc,:account,
              :subtotal,:vat,:total,:status,'غير مرحّل',:notes,:uid,:created)""",
              {"supplier_no":request.form["supplier_invoice_no"],"internal":internal,
               "dt":request.form["invoice_date"],"supplier":request.form["supplier_id"],
               "po":request.form.get("po_id") or None,"grn":request.form.get("grn_id") or None,
               "cc":request.form.get("cost_center_id") or None,
               "account":request.form["expense_or_inventory_account_id"],
               "subtotal":subtotal,"vat":vat,"total":total,
               "status":request.form.get("status","مسودة"),
               "notes":request.form.get("notes",""),
               "uid":session.get("user_id"),"created":datetime.now()})
        except Exception:
            db.session.rollback()
            flash("رقم فاتورة المورد مستخدم مسبقًا لهذا المورد","danger")
            return redirect(url_for("supplier_invoices"))
        inv_id=row("SELECT id FROM supplier_invoices WHERE internal_no=:n",{"n":internal})["id"]
        if request.form.get("status")=="معتمدة":
            try:
                post_supplier_invoice(inv_id)
                flash(f"تم حفظ وترحيل فاتورة المورد {internal}","success")
            except Exception as exc:
                db.session.rollback()
                execute("UPDATE supplier_invoices SET status='مسودة',posting_status='خطأ' WHERE id=:id",{"id":inv_id})
                flash(f"تم حفظها كمسودة ولم تُرحّل: {exc}","danger")
        else:flash(f"تم حفظ فاتورة المورد {internal} كمسودة","success")
        return redirect(url_for("supplier_invoice_view",invoice_id=inv_id))
    accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                     WHERE active=1 AND accepts_entries=1 ORDER BY account_code""")
    return render_template("supplier_invoices.html",
      invoices=rows("""SELECT si.*,s.name supplier_name,j.journal_no FROM supplier_invoices si
                       JOIN suppliers s ON s.id=si.supplier_id
                       LEFT JOIN journal_entries j ON j.id=si.journal_id
                       ORDER BY si.invoice_date DESC,si.id DESC"""),
      suppliers=rows("SELECT id,name,name_en FROM suppliers ORDER BY name"),
      pos=rows("SELECT id,po_no,supplier_id,total FROM purchase_orders ORDER BY po_date DESC"),
      grns=rows("SELECT id,grn_no,supplier_id,po_id FROM goods_receipts ORDER BY grn_date DESC"),
      centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"),
      accounts=accounts)

@app.route("/supplier-invoices/<int:invoice_id>")
@login_required
def supplier_invoice_view(invoice_id):
    inv=row("""SELECT si.*,s.name supplier_name,s.name_en supplier_name_en,
              s.vat_number supplier_vat,po.po_no,g.grn_no,a.account_code,a.account_name_ar,
              j.journal_no FROM supplier_invoices si
              JOIN suppliers s ON s.id=si.supplier_id
              LEFT JOIN purchase_orders po ON po.id=si.po_id
              LEFT JOIN goods_receipts g ON g.id=si.grn_id
              JOIN chart_of_accounts a ON a.id=si.expense_or_inventory_account_id
              LEFT JOIN journal_entries j ON j.id=si.journal_id WHERE si.id=:id""",{"id":invoice_id})
    if not inv:return "فاتورة المورد غير موجودة",404
    return render_template("supplier_invoice_view.html",inv=inv)

@app.route("/supplier-invoices/<int:invoice_id>/post",methods=["POST"])
@login_required
def supplier_invoice_post(invoice_id):
    try:
        post_supplier_invoice(invoice_id)
        flash("تم ترحيل فاتورة المورد","success")
    except Exception as exc:flash(str(exc),"danger")
    return redirect(url_for("supplier_invoice_view",invoice_id=invoice_id))




@app.route("/sales/tracking")
@login_required
def sales_tracking():
    query=request.args.get("q","").strip()
    conditions=[];params={}
    if query:
        conditions.append("""(q.quotation_no ILIKE :q OR o.order_no ILIKE :q
                           OR d.delivery_no ILIKE :q OR i.invoice_no ILIKE :q
                           OR c.name ILIKE :q)""")
        params["q"]=f"%{query}%"
    where=" WHERE "+" AND ".join(conditions) if conditions else ""
    docs=rows(f"""SELECT
      q.id quotation_id,q.quotation_no,q.status quotation_status,
      o.id order_id,o.order_no,o.status order_status,
      d.id delivery_id,d.delivery_no,d.status delivery_status,
      i.id invoice_id,i.invoice_no,i.status invoice_status,i.payment_status,
      c.name customer_name,
      COALESCE(i.total,o.total,q.total,0) total,
      COALESCE(i.invoice_date,d.delivery_date,o.order_date,q.quotation_date) document_date,
      COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                WHERE a.invoice_id=i.id),0) paid,
      COALESCE(i.total,0)-COALESCE((SELECT SUM(a.allocated_amount)
                FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0) outstanding
      FROM sales_quotations q
      LEFT JOIN sales_orders o ON o.quotation_id=q.id
      LEFT JOIN sales_deliveries d ON d.order_id=o.id
      LEFT JOIN invoices i ON i.delivery_id=d.id
      JOIN customers c ON c.id=q.customer_id
      {where}
      ORDER BY document_date DESC,q.id DESC""",params)

    direct_orders=rows(f"""SELECT
      NULL quotation_id,NULL quotation_no,NULL quotation_status,
      o.id order_id,o.order_no,o.status order_status,
      d.id delivery_id,d.delivery_no,d.status delivery_status,
      i.id invoice_id,i.invoice_no,i.status invoice_status,i.payment_status,
      c.name customer_name,
      COALESCE(i.total,o.total,0) total,
      COALESCE(i.invoice_date,d.delivery_date,o.order_date) document_date,
      COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                WHERE a.invoice_id=i.id),0) paid,
      COALESCE(i.total,0)-COALESCE((SELECT SUM(a.allocated_amount)
                FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0) outstanding
      FROM sales_orders o
      LEFT JOIN sales_deliveries d ON d.order_id=o.id
      LEFT JOIN invoices i ON i.delivery_id=d.id
      JOIN customers c ON c.id=o.customer_id
      WHERE o.quotation_id IS NULL
      ORDER BY document_date DESC,o.id DESC""")
    return render_template("sales_tracking.html",docs=list(docs)+list(direct_orders),q=query)

@app.route("/sales/timeline")
@login_required
def sales_timeline():
    invoice_id=request.args.get("invoice_id",type=int)
    delivery_id=request.args.get("delivery_id",type=int)
    order_id=request.args.get("order_id",type=int)
    quotation_id=request.args.get("quotation_id",type=int)
    timeline=sales_document_timeline(invoice_id,delivery_id,order_id,quotation_id)
    return render_template("sales_timeline.html",timeline=timeline)

@app.route("/invoices/<int:invoice_id>/commercial-data",methods=["GET","POST"])
@login_required
def invoice_commercial_data(invoice_id):
    inv=row("""SELECT i.*,c.name customer_name FROM invoices i
               JOIN customers c ON c.id=i.customer_id WHERE i.id=:id""",{"id":invoice_id})
    if not inv:
        return "الفاتورة غير موجودة",404
    if request.method=="POST":
        due_date=request.form.get("due_date") or None
        execute("""UPDATE invoices SET due_date=:due,payment_terms=:terms,
          payment_method=:method,sales_person=:sales_person,
          customer_reference=:customer_reference,notes=:notes
          WHERE id=:id""",
          {"due":due_date,"terms":request.form.get("payment_terms",""),
           "method":request.form.get("payment_method",""),
           "sales_person":request.form.get("sales_person",""),
           "customer_reference":request.form.get("customer_reference",""),
           "notes":request.form.get("notes",""),"id":invoice_id})
        audit("UPDATE","INVOICE",f"تحديث البيانات التجارية للفاتورة {inv['invoice_no']}")
        flash("تم تحديث بيانات الفاتورة","success")
        return redirect(url_for("invoice_view",invoice_id=invoice_id))
    return render_template("invoice_commercial_data.html",invoice=inv)


@app.route("/receivables",methods=["GET","POST"])
@login_required
def receivables():
    if request.method=="POST":
        customer_id=int(request.form["customer_id"])
        receipt_date=request.form["receipt_date"]
        cash_account_id=int(request.form["cash_bank_account_id"])
        invoice_ids=request.form.getlist("invoice_id[]")
        amounts=request.form.getlist("allocated_amount[]")
        allocations=[]
        for idx,invoice_id in enumerate(invoice_ids):
            amount=round(float(amounts[idx] or 0),2)
            if amount>0:
                allocations.append({"invoice_id":int(invoice_id),"amount":amount})
        total=round(sum(x["amount"] for x in allocations),2)
        if total<=0:
            flash("أدخل مبلغ سداد لفاتورة واحدة على الأقل","danger")
            return redirect(url_for("receivables"))

        settings_row=row("SELECT customer_account_id FROM settings WHERE id=1")
        if not settings_row or not settings_row["customer_account_id"]:
            flash("حدد حساب العملاء الافتراضي من الإعدادات أولًا","danger")
            return redirect(url_for("receivables"))

        try:
            ensure_open_period(receipt_date)
            voucher_no=next_treasury_number("قبض",receipt_date)
            execute("""INSERT INTO treasury_vouchers(
                voucher_no,voucher_type,voucher_date,party_type,customer_id,
                cash_bank_account_id,counter_account_id,amount,payment_method,
                reference,description,status,posting_status,created_by,created_at)
                VALUES(:no,'قبض',:dt,'عميل',:customer,:cash,:counter,:amount,:method,
                :reference,:description,'معتمد','غير مرحّل',:uid,:created)""",
                {"no":voucher_no,"dt":receipt_date,"customer":customer_id,
                 "cash":cash_account_id,"counter":settings_row["customer_account_id"],
                 "amount":total,"method":request.form.get("payment_method","تحويل بنكي"),
                 "reference":request.form.get("reference",""),
                 "description":request.form.get("description","تحصيل فواتير عميل"),
                 "uid":session.get("user_id"),"created":datetime.now()})
            voucher_id=row("SELECT id FROM treasury_vouchers WHERE voucher_no=:no",
                           {"no":voucher_no})["id"]
            post_treasury_voucher(voucher_id)
            allocate_customer_receipt(voucher_id,allocations)
            audit("CREATE","CUSTOMER_RECEIPT",f"تحصيل وربط الفواتير بسند {voucher_no}")
            flash(f"تم إنشاء وترحيل سند القبض {voucher_no} وربطه بالفواتير","success")
            return redirect(url_for("treasury_view",voucher_id=voucher_id))
        except Exception as exc:
            db.session.rollback()
            flash(str(exc),"danger")
            return redirect(url_for("receivables"))

    open_invoices=rows("""SELECT i.id,i.invoice_no,i.invoice_date,
      COALESCE(i.due_date,i.invoice_date) due_date,i.customer_id,c.name customer_name,
      i.total,COALESCE((SELECT SUM(a.allocated_amount)
        FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0) paid,
      i.total-COALESCE((SELECT SUM(a.allocated_amount)
        FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0) outstanding,
      i.payment_status
      FROM invoices i JOIN customers c ON c.id=i.customer_id
      WHERE i.status='معتمدة'
        AND i.total-COALESCE((SELECT SUM(a.allocated_amount)
          FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0)>0.005
      ORDER BY c.name,i.invoice_date,i.id""")
    accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                     WHERE active=1 AND accepts_entries=1
                     AND account_type='أصل' ORDER BY account_code""")
    return render_template("receivables.html",open_invoices=open_invoices,accounts=accounts,
      customers=rows("SELECT id,name,name_en FROM customers ORDER BY name"))

@app.route("/receivables/customer/<int:customer_id>/open-invoices")
@login_required
def customer_open_invoices(customer_id):
    data=rows("""SELECT i.id,i.invoice_no,i.invoice_date,
      COALESCE(i.due_date,i.invoice_date) due_date,i.total,
      COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                WHERE a.invoice_id=i.id),0) paid,
      i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                WHERE a.invoice_id=i.id),0) outstanding
      FROM invoices i WHERE i.customer_id=:customer AND i.status='معتمدة'
      AND i.total-COALESCE((SELECT SUM(a.allocated_amount)
                FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0)>0.005
      ORDER BY i.invoice_date,i.id""",{"customer":customer_id})
    return {"items":[dict(x) for x in data]}

@app.route("/receivables/aging")
@login_required
def receivables_aging():
    data=rows("""SELECT c.id,c.name,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date)<=30 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) bucket_0_30,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date) BETWEEN 31 AND 60 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) bucket_31_60,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date) BETWEEN 61 AND 90 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) bucket_61_90,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date)>90 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) bucket_over_90,
      SUM(i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0)) total_outstanding
      FROM invoices i JOIN customers c ON c.id=i.customer_id
      WHERE i.status='معتمدة'
      AND i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                            WHERE a.invoice_id=i.id),0)>0.005
      GROUP BY c.id,c.name ORDER BY total_outstanding DESC""")
    return render_template("receivables_aging.html",rows=data)

@app.route("/receivables/aging.xlsx")
@login_required
def receivables_aging_export():
    data=rows("""SELECT c.name,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date)<=30 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) b1,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date) BETWEEN 31 AND 60 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) b2,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date) BETWEEN 61 AND 90 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) b3,
      SUM(CASE WHEN CURRENT_DATE-COALESCE(i.due_date,i.invoice_date)>90 THEN
        i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                          WHERE a.invoice_id=i.id),0) ELSE 0 END) b4,
      SUM(i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                           WHERE a.invoice_id=i.id),0)) total
      FROM invoices i JOIN customers c ON c.id=i.customer_id
      WHERE i.status='معتمدة'
      AND i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                            WHERE a.invoice_id=i.id),0)>0.005
      GROUP BY c.id,c.name ORDER BY total DESC""")
    return xlsx_response("receivables_aging.xlsx","أعمار ديون العملاء",
      ["العميل","0-30 يوم","31-60 يوم","61-90 يوم","أكثر من 90 يوم","الإجمالي"],
      [[r["name"],r["b1"],r["b2"],r["b3"],r["b4"],r["total"]] for r in data])


@app.route("/treasury", methods=["GET","POST"])
@login_required
def treasury():
    if request.method == "POST":
        voucher_type = request.form["voucher_type"]
        voucher_date = request.form["voucher_date"]
        amount = round(float(request.form["amount"] or 0),2)
        if amount <= 0:
            flash("المبلغ يجب أن يكون أكبر من صفر","danger")
            return redirect(url_for("treasury"))

        try:
            ensure_open_period(voucher_date)
        except Exception as exc:
            flash(str(exc),"danger")
            return redirect(url_for("treasury"))

        voucher_no = next_treasury_number(voucher_type, voucher_date)
        execute("""INSERT INTO treasury_vouchers(
            voucher_no,voucher_type,voucher_date,party_type,customer_id,supplier_id,
            cash_bank_account_id,counter_account_id,amount,payment_method,reference,
            description,cost_center_id,status,posting_status,created_by,created_at)
            VALUES(:no,:type,:dt,:ptype,:cid,:sid,:cash,:counter,:amount,:method,
                   :ref,:des,:cc,:status,'غير مرحّل',:uid,:created)""",
            {"no":voucher_no,"type":voucher_type,"dt":voucher_date,
             "ptype":request.form.get("party_type",""),
             "cid":request.form.get("customer_id") or None,
             "sid":request.form.get("supplier_id") or None,
             "cash":request.form["cash_bank_account_id"],
             "counter":request.form["counter_account_id"],
             "amount":amount,"method":request.form.get("payment_method","نقدي"),
             "ref":request.form.get("reference",""),
             "des":request.form.get("description",""),
             "cc":request.form.get("cost_center_id") or None,
             "status":request.form.get("status","مسودة"),
             "uid":session.get("user_id"),"created":datetime.now()})
        voucher_id = row("SELECT id FROM treasury_vouchers WHERE voucher_no=:n",{"n":voucher_no})["id"]

        if request.form.get("status") == "معتمد":
            try:
                post_treasury_voucher(voucher_id)
                flash(f"تم حفظ وترحيل السند {voucher_no}","success")
            except Exception as exc:
                db.session.rollback()
                execute("UPDATE treasury_vouchers SET status='مسودة',posting_status='خطأ' WHERE id=:id",
                        {"id":voucher_id})
                flash(f"تم حفظ السند كمسودة ولم يُرحّل: {exc}","danger")
        else:
            flash(f"تم حفظ السند {voucher_no} كمسودة","success")

        if request.form.get("print_after_save") == "1":
            return redirect(url_for("treasury_view",voucher_id=voucher_id,print=1))
        return redirect(url_for("treasury_view",voucher_id=voucher_id))

    q = request.args.get("q","").strip()
    conditions=[]; params={}
    if q:
        conditions.append("""(tv.voucher_no ILIKE :q OR COALESCE(tv.reference,'') ILIKE :q
                           OR COALESCE(tv.description,'') ILIKE :q
                           OR COALESCE(c.name,'') ILIKE :q OR COALESCE(s.name,'') ILIKE :q)""")
        params["q"]=f"%{q}%"
    where = " WHERE " + " AND ".join(conditions) if conditions else ""
    vouchers = rows(f"""SELECT tv.*,c.name customer_name,s.name supplier_name,
                        a.account_name_ar cash_account_name,b.account_name_ar counter_account_name,
                        j.journal_no
                        FROM treasury_vouchers tv
                        LEFT JOIN customers c ON c.id=tv.customer_id
                        LEFT JOIN suppliers s ON s.id=tv.supplier_id
                        JOIN chart_of_accounts a ON a.id=tv.cash_bank_account_id
                        JOIN chart_of_accounts b ON b.id=tv.counter_account_id
                        LEFT JOIN journal_entries j ON j.id=tv.journal_id
                        {where}
                        ORDER BY tv.voucher_date DESC,tv.id DESC""",params)
    accounts = rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                       WHERE active=1 AND accepts_entries=1 ORDER BY account_code""")
    return render_template("treasury.html",vouchers=vouchers,accounts=accounts,
        customers=rows("SELECT id,name,name_en,vat_number FROM customers ORDER BY name"),
        suppliers=rows("SELECT id,name,name_en,vat_number FROM suppliers ORDER BY name"),
        centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"),
        q=q)

@app.route("/treasury/<int:voucher_id>")
@login_required
def treasury_view(voucher_id):
    voucher = row("""SELECT tv.*,c.name customer_name,c.name_en customer_name_en,
                     s.name supplier_name,s.name_en supplier_name_en,
                     a.account_code cash_account_code,a.account_name_ar cash_account_name,
                     b.account_code counter_account_code,b.account_name_ar counter_account_name,
                     cc.code cost_center_code,cc.name cost_center_name,j.journal_no
                     FROM treasury_vouchers tv
                     LEFT JOIN customers c ON c.id=tv.customer_id
                     LEFT JOIN suppliers s ON s.id=tv.supplier_id
                     JOIN chart_of_accounts a ON a.id=tv.cash_bank_account_id
                     JOIN chart_of_accounts b ON b.id=tv.counter_account_id
                     LEFT JOIN cost_centers cc ON cc.id=tv.cost_center_id
                     LEFT JOIN journal_entries j ON j.id=tv.journal_id
                     WHERE tv.id=:id""",{"id":voucher_id})
    if not voucher:
        return "السند غير موجود",404
    company=row("SELECT * FROM settings WHERE id=1")
    return render_template("treasury_view.html",voucher=voucher,company=company,
                           print_lang=request.args.get("lang","ar"))

@app.route("/treasury/<int:voucher_id>/post", methods=["POST"])
@login_required
def treasury_post(voucher_id):
    try:
        post_treasury_voucher(voucher_id)
        flash("تم ترحيل السند محاسبيًا","success")
    except Exception as exc:
        flash(str(exc),"danger")
    return redirect(url_for("treasury_view",voucher_id=voucher_id))

@app.route("/treasury/<int:voucher_id>/delete", methods=["POST"])
@login_required
def treasury_delete(voucher_id):
    v=row("SELECT * FROM treasury_vouchers WHERE id=:id",{"id":voucher_id})
    if not v:
        return "السند غير موجود",404
    if v["journal_id"]:
        flash("لا يمكن حذف سند مرحّل. يجب إنشاء قيد عكسي.","danger")
    else:
        execute("DELETE FROM treasury_vouchers WHERE id=:id",{"id":voucher_id})
        audit("DELETE","TREASURY",f"حذف السند {v['voucher_no']}")
        flash("تم حذف السند","success")
    return redirect(url_for("treasury"))

@app.route("/treasury/export.xlsx")
@login_required
def treasury_export():
    data=rows("""SELECT voucher_no,voucher_type,voucher_date,party_type,amount,
                 payment_method,reference,description,status,posting_status
                 FROM treasury_vouchers ORDER BY voucher_date DESC,id DESC""")
    headers=["رقم السند","النوع","التاريخ","نوع الطرف","المبلغ","طريقة الدفع",
             "المرجع","البيان","الحالة","الترحيل"]
    keys=["voucher_no","voucher_type","voucher_date","party_type","amount",
          "payment_method","reference","description","status","posting_status"]
    return xlsx_response("treasury_vouchers.xlsx","سندات الخزينة",headers,
                         [[r.get(k) for k in keys] for r in data])


@app.route("/fiscal-periods",methods=["GET","POST"])
@login_required
def fiscal_periods():
    if request.method=="POST":
        pid=request.form.get("period_id",type=int); action=request.form.get("action")
        if pid and action in ("open","close"):
            status="مفتوحة" if action=="open" else "مغلقة"
            execute("UPDATE fiscal_periods SET status=:s WHERE id=:id",{"s":status,"id":pid})
            audit("PERIOD","FISCAL_PERIOD",f"الفترة {pid}: {status}")
        return redirect(url_for("fiscal_periods"))
    return render_template("fiscal_periods.html",periods=rows("""SELECT p.*,y.name fiscal_year_name
      FROM fiscal_periods p JOIN fiscal_years y ON y.id=p.fiscal_year_id ORDER BY p.start_date DESC"""))

@app.route("/invoices/<int:invoice_id>/post",methods=["POST"])
@login_required
def invoice_post(invoice_id):
    try:
        post_invoice_to_ledger(invoice_id)
        execute("UPDATE invoices SET status='معتمدة' WHERE id=:id",{"id":invoice_id})
        flash("تم ترحيل الفاتورة محاسبيًا","success")
    except Exception as exc: flash(str(exc),"danger")
    return redirect(url_for("invoice_view",invoice_id=invoice_id))

@app.route("/expenses/<int:expense_id>/post",methods=["POST"])
@login_required
def expense_post(expense_id):
    try:
        post_expense_to_ledger(expense_id)
        execute("UPDATE expenses SET status='معتمدة' WHERE id=:id",{"id":expense_id})
        flash("تم ترحيل المصروف محاسبيًا","success")
    except Exception as exc: flash(str(exc),"danger")
    return redirect(url_for("expenses"))




@app.route("/notifications")
@login_required
def notifications_center():
    notifications=[]
    for r in rows("""SELECT id,sku,name,quantity,reorder_level FROM inventory
                     WHERE active=1 AND quantity<=reorder_level
                     ORDER BY quantity,name LIMIT 100"""):
        notifications.append({"level":"warning","title":"مخزون منخفض",
          "message":f"{r['sku'] or ''} - {r['name']}، الرصيد {r['quantity']} وحد إعادة الطلب {r['reorder_level']}",
          "url":url_for("inventory_item_view",item_id=r["id"])})

    for r in rows("""SELECT i.id,i.invoice_no,c.name customer_name,
      i.total-COALESCE((SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
                        WHERE a.invoice_id=i.id),0) outstanding
      FROM invoices i JOIN customers c ON c.id=i.customer_id
      WHERE i.status='معتمدة' AND COALESCE(i.due_date,i.invoice_date)<CURRENT_DATE
      AND i.total-COALESCE((SELECT SUM(a.allocated_amount)
                           FROM invoice_payment_allocations a WHERE a.invoice_id=i.id),0)>0.005
      ORDER BY COALESCE(i.due_date,i.invoice_date) LIMIT 100"""):
        notifications.append({"level":"danger","title":"فاتورة عميل متأخرة",
          "message":f"{r['invoice_no']} - {r['customer_name']}، المتبقي {float(r['outstanding']):.2f}",
          "url":url_for("invoice_view",invoice_id=r["id"])})

    for r in rows("""SELECT id,journal_no,description FROM journal_entries
                     WHERE status<>'مرحّل' ORDER BY journal_date,id LIMIT 100"""):
        notifications.append({"level":"info","title":"قيد غير مرحّل",
          "message":f"{r['journal_no']} - {r['description'] or ''}",
          "url":url_for("journal_view",journal_id=r["id"])})

    return render_template("notifications_center.html",notifications=notifications)

@app.route("/system-health")
@login_required
def system_health():
    checks=[]
    try:
        db.session.execute(text("SELECT 1"))
        checks.append({"name":"قاعدة البيانات","status":"سليم","ok":True})
    except Exception as exc:
        checks.append({"name":"قاعدة البيانات","status":str(exc),"ok":False})
    checks += [
        {"name":"إصدار النظام","status":APP_VERSION,"ok":True},
        {"name":"المستخدمون","status":row("SELECT COUNT(*) c FROM users")["c"],"ok":True},
        {"name":"القيود","status":row("SELECT COUNT(*) c FROM journal_entries")["c"],"ok":True},
        {"name":"الفواتير","status":row("SELECT COUNT(*) c FROM invoices")["c"],"ok":True},
        {"name":"الأصناف","status":row("SELECT COUNT(*) c FROM inventory")["c"],"ok":True},
    ]
    return render_template("system_health.html",checks=checks)




@app.route("/hr-payroll")
@login_required
def hr_payroll_center():
    stats={
      "employees":row("SELECT COUNT(*) c FROM employees WHERE active=1")["c"],
      "departments":row("SELECT COUNT(*) c FROM departments WHERE active=1")["c"],
      "runs":row("SELECT COUNT(*) c FROM payroll_runs")["c"],
      "net":row("SELECT COALESCE(SUM(total_net),0) v FROM payroll_runs WHERE status='مرحّل'")["v"],
    }
    return render_template("hr_payroll_center.html",stats=stats)

@app.route("/departments",methods=["GET","POST"])
@login_required
def departments():
    if request.method=="POST":
        execute("""INSERT INTO departments(code,name,active)
                   VALUES(:code,:name,1)
                   ON CONFLICT(code) DO UPDATE SET name=EXCLUDED.name""",
                {"code":request.form["code"],"name":request.form["name"]})
        flash("تم حفظ القسم","success")
        return redirect(url_for("departments"))
    return render_template("departments.html",
      rows=rows("SELECT * FROM departments ORDER BY code"))

@app.route("/payroll/settings",methods=["GET","POST"])
@login_required
def payroll_settings():
    cfg=payroll_configuration()
    if request.method=="POST":
        execute("""UPDATE payroll_settings SET
          salary_expense_account_id=:expense,payroll_payable_account_id=:payable,
          deduction_account_id=:deduction,working_days_per_month=:days,
          working_hours_per_day=:hours,overtime_multiplier=:multiplier WHERE id=1""",
          {"expense":request.form.get("salary_expense_account_id") or None,
           "payable":request.form.get("payroll_payable_account_id") or None,
           "deduction":request.form.get("deduction_account_id") or None,
           "days":float(request.form.get("working_days_per_month") or 30),
           "hours":float(request.form.get("working_hours_per_day") or 8),
           "multiplier":float(request.form.get("overtime_multiplier") or 1.5)})
        flash("تم حفظ إعدادات الرواتب","success")
        return redirect(url_for("payroll_settings"))
    return render_template("payroll_settings.html",config=cfg,
      accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                       WHERE active=1 AND accepts_entries=1 ORDER BY account_code"""))

@app.route("/attendance",methods=["GET","POST"])
@login_required
def attendance():
    if request.method=="POST":
        execute("""INSERT INTO attendance_records(employee_id,attendance_date,status,
          check_in,check_out,overtime_hours,absence_hours,notes,created_by,created_at)
          VALUES(:employee,:dt,:status,:check_in,:check_out,:ot,:absence,:notes,:uid,:created)
          ON CONFLICT(employee_id,attendance_date) DO UPDATE SET
          status=EXCLUDED.status,check_in=EXCLUDED.check_in,check_out=EXCLUDED.check_out,
          overtime_hours=EXCLUDED.overtime_hours,absence_hours=EXCLUDED.absence_hours,
          notes=EXCLUDED.notes""",
          {"employee":request.form["employee_id"],"dt":request.form["attendance_date"],
           "status":request.form.get("status","حاضر"),
           "check_in":request.form.get("check_in") or None,
           "check_out":request.form.get("check_out") or None,
           "ot":float(request.form.get("overtime_hours") or 0),
           "absence":float(request.form.get("absence_hours") or 0),
           "notes":request.form.get("notes",""),"uid":session.get("user_id"),
           "created":datetime.now()})
        flash("تم حفظ سجل الحضور","success")
        return redirect(url_for("attendance"))
    date_from=request.args.get("date_from","")
    date_to=request.args.get("date_to","")
    conditions=[];params={}
    if date_from:
        conditions.append("a.attendance_date>=:date_from");params["date_from"]=date_from
    if date_to:
        conditions.append("a.attendance_date<=:date_to");params["date_to"]=date_to
    where=" WHERE "+" AND ".join(conditions) if conditions else ""
    return render_template("attendance.html",date_from=date_from,date_to=date_to,
      records=rows(f"""SELECT a.*,e.employee_no,e.name FROM attendance_records a
                       JOIN employees e ON e.id=a.employee_id
                       {where} ORDER BY a.attendance_date DESC,e.name""",params),
      employees=rows("SELECT id,employee_no,name FROM employees WHERE active=1 ORDER BY name"))

@app.route("/salary-adjustments",methods=["GET","POST"])
@login_required
def salary_adjustments():
    if request.method=="POST":
        execute("""INSERT INTO employee_salary_adjustments(employee_id,adjustment_date,
          adjustment_type,amount,recurring,description,active,created_by,created_at)
          VALUES(:employee,:dt,:type,:amount,:recurring,:description,1,:uid,:created)""",
          {"employee":request.form["employee_id"],"dt":request.form["adjustment_date"],
           "type":request.form["adjustment_type"],
           "amount":float(request.form["amount"] or 0),
           "recurring":1 if request.form.get("recurring") else 0,
           "description":request.form.get("description",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        flash("تم حفظ الاستحقاق/الاستقطاع","success")
        return redirect(url_for("salary_adjustments"))
    return render_template("salary_adjustments.html",
      rows=rows("""SELECT a.*,e.name employee_name FROM employee_salary_adjustments a
                   JOIN employees e ON e.id=a.employee_id ORDER BY a.adjustment_date DESC,a.id DESC"""),
      employees=rows("SELECT id,employee_no,name FROM employees WHERE active=1 ORDER BY name"))

@app.route("/payroll/runs",methods=["GET","POST"])
@login_required
def payroll_runs():
    if request.method=="POST":
        start=request.form["period_start"];end=request.form["period_end"]
        no=next_payroll_run_number(end)
        execute("""INSERT INTO payroll_runs(run_no,period_start,period_end,payment_date,
          status,notes,created_by,created_at)
          VALUES(:no,:start,:end,:payment,'مسودة',:notes,:uid,:created)""",
          {"no":no,"start":start,"end":end,
           "payment":request.form.get("payment_date") or end,
           "notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        run_id=row("SELECT id FROM payroll_runs WHERE run_no=:no",{"no":no})["id"]
        total_gross=total_deductions=total_net=0
        employees_data=rows("""SELECT e.*,COALESCE(e.cost_center_id,cc.id) resolved_cc
          FROM employees e LEFT JOIN cost_centers cc ON cc.id=e.cost_center_id
          WHERE e.active=1 ORDER BY e.id""")
        for employee in employees_data:
            calc=calculate_employee_payroll(employee,start,end)
            execute("""INSERT INTO payroll_lines(run_id,employee_id,basic_salary,allowances,
              overtime_hours,overtime_amount,absence_days,absence_deduction,other_deductions,
              gross_salary,total_deductions,net_salary,cost_center_id)
              VALUES(:run,:employee,:basic,:allowances,:ot_hours,:ot_amount,:absence_days,
              :absence_deduction,:other_deductions,:gross,:deductions,:net,:cc)""",
              {"run":run_id,"employee":employee["id"],"basic":calc["basic_salary"],
               "allowances":calc["allowances"],"ot_hours":calc["overtime_hours"],
               "ot_amount":calc["overtime_amount"],"absence_days":calc["absence_days"],
               "absence_deduction":calc["absence_deduction"],
               "other_deductions":calc["other_deductions"],"gross":calc["gross_salary"],
               "deductions":calc["total_deductions"],"net":calc["net_salary"],
               "cc":employee.get("cost_center_id")})
            total_gross+=calc["gross_salary"]
            total_deductions+=calc["total_deductions"]
            total_net+=calc["net_salary"]
        execute("""UPDATE payroll_runs SET total_gross=:gross,total_deductions=:deductions,
                   total_net=:net WHERE id=:id""",
                {"gross":round(total_gross,2),"deductions":round(total_deductions,2),
                 "net":round(total_net,2),"id":run_id})
        flash(f"تم إنشاء مسير الرواتب {no}","success")
        return redirect(url_for("payroll_run_view",run_id=run_id))
    return render_template("payroll_runs.html",
      runs=rows("""SELECT r.*,j.journal_no FROM payroll_runs r
                   LEFT JOIN journal_entries j ON j.id=r.journal_id
                   ORDER BY r.period_end DESC,r.id DESC"""))

@app.route("/payroll/runs/<int:run_id>")
@login_required
def payroll_run_view(run_id):
    run=row("""SELECT r.*,j.journal_no FROM payroll_runs r
               LEFT JOIN journal_entries j ON j.id=r.journal_id WHERE r.id=:id""",{"id":run_id})
    if not run:
        return "مسير الرواتب غير موجود",404
    lines=rows("""SELECT pl.*,e.employee_no,e.name,e.job_title,b.name branch_name
      FROM payroll_lines pl JOIN employees e ON e.id=pl.employee_id
      LEFT JOIN branches b ON b.id=e.branch_id WHERE pl.run_id=:id ORDER BY e.name""",{"id":run_id})
    return render_template("payroll_run_view.html",run=run,lines=lines)

@app.route("/payroll/runs/<int:run_id>/post",methods=["POST"])
@login_required
def payroll_run_post(run_id):
    try:
        jid=post_payroll_run(run_id)
        flash("تم اعتماد وترحيل مسير الرواتب","success")
    except Exception as exc:
        db.session.rollback()
        flash(str(exc),"danger")
    return redirect(url_for("payroll_run_view",run_id=run_id))

@app.route("/payroll/payslip/<int:line_id>")
@login_required
def payroll_payslip(line_id):
    line=row("""SELECT pl.*,pr.run_no,pr.period_start,pr.period_end,pr.payment_date,
      e.employee_no,e.name,e.name_en,e.job_title,e.bank_iban,b.name branch_name
      FROM payroll_lines pl JOIN payroll_runs pr ON pr.id=pl.run_id
      JOIN employees e ON e.id=pl.employee_id
      LEFT JOIN branches b ON b.id=e.branch_id WHERE pl.id=:id""",{"id":line_id})
    if not line:
        return "كشف الراتب غير موجود",404
    return render_template("payroll_payslip.html",line=line)

@app.route("/payroll/report")
@login_required
def payroll_report():
    data=rows("""SELECT pr.run_no,pr.period_start,pr.period_end,e.employee_no,e.name,
      pl.basic_salary,pl.allowances,pl.overtime_amount,pl.gross_salary,
      pl.total_deductions,pl.net_salary
      FROM payroll_lines pl JOIN payroll_runs pr ON pr.id=pl.run_id
      JOIN employees e ON e.id=pl.employee_id
      ORDER BY pr.period_end DESC,e.name""")
    return render_template("payroll_report.html",rows=data)

@app.route("/payroll/report.xlsx")
@login_required
def payroll_report_export():
    data=rows("""SELECT pr.run_no,pr.period_start,pr.period_end,e.employee_no,e.name,
      pl.basic_salary,pl.allowances,pl.overtime_amount,pl.gross_salary,
      pl.total_deductions,pl.net_salary
      FROM payroll_lines pl JOIN payroll_runs pr ON pr.id=pl.run_id
      JOIN employees e ON e.id=pl.employee_id
      ORDER BY pr.period_end DESC,e.name""")
    return xlsx_response("payroll_report.xlsx","الرواتب",
      ["المسير","من","إلى","رقم الموظف","الموظف","الأساسي","البدلات",
       "الإضافي","الإجمالي","الاستقطاعات","الصافي"],
      [[r["run_no"],r["period_start"],r["period_end"],r["employee_no"],r["name"],
        r["basic_salary"],r["allowances"],r["overtime_amount"],r["gross_salary"],
        r["total_deductions"],r["net_salary"]] for r in data])


@app.route("/fixed-assets")
@login_required
def fixed_assets_center():
    stats={
      "assets":row("SELECT COUNT(*) c FROM fixed_assets")["c"],
      "active":row("SELECT COUNT(*) c FROM fixed_assets WHERE status='نشط'")["c"],
      "cost":row("SELECT COALESCE(SUM(purchase_cost),0) v FROM fixed_assets")["v"],
      "nbv":row("SELECT COALESCE(SUM(net_book_value),0) v FROM fixed_assets")["v"],
    }
    return render_template("fixed_assets_center.html",stats=stats)

@app.route("/fixed-assets/categories",methods=["GET","POST"])
@login_required
def asset_categories():
    if request.method=="POST":
        execute("""INSERT INTO asset_categories(code,name,useful_life_months,depreciation_method,
          asset_account_id,accumulated_depr_account_id,depreciation_expense_account_id,active)
          VALUES(:code,:name,:life,:method,:asset,:accum,:expense,1)
          ON CONFLICT(code) DO UPDATE SET name=EXCLUDED.name""",
          {"code":request.form["code"],"name":request.form["name"],
           "life":int(request.form.get("useful_life_months") or 60),
           "method":request.form.get("depreciation_method","القسط الثابت"),
           "asset":request.form.get("asset_account_id") or None,
           "accum":request.form.get("accumulated_depr_account_id") or None,
           "expense":request.form.get("depreciation_expense_account_id") or None})
        flash("تم حفظ فئة الأصل","success")
        return redirect(url_for("asset_categories"))
    return render_template("asset_categories.html",
      rows=rows("""SELECT ac.*,a1.account_name_ar asset_account,
        a2.account_name_ar accumulated_account,a3.account_name_ar expense_account
        FROM asset_categories ac
        LEFT JOIN chart_of_accounts a1 ON a1.id=ac.asset_account_id
        LEFT JOIN chart_of_accounts a2 ON a2.id=ac.accumulated_depr_account_id
        LEFT JOIN chart_of_accounts a3 ON a3.id=ac.depreciation_expense_account_id
        ORDER BY ac.code"""),
      accounts=rows("""SELECT id,account_code,account_name_ar FROM chart_of_accounts
                       WHERE active=1 AND accepts_entries=1 ORDER BY account_code"""))

@app.route("/fixed-assets/assets",methods=["GET","POST"])
@login_required
def fixed_assets():
    if request.method=="POST":
        category=row("SELECT * FROM asset_categories WHERE id=:id",{"id":request.form["category_id"]})
        no=next_asset_number()
        cost=float(request.form.get("purchase_cost") or 0)
        residual=float(request.form.get("residual_value") or 0)
        execute("""INSERT INTO fixed_assets(asset_no,name,name_en,category_id,branch_id,cost_center_id,
          supplier_id,purchase_date,capitalization_date,purchase_cost,residual_value,
          useful_life_months,depreciation_method,accumulated_depreciation,net_book_value,
          serial_number,location,status,notes,created_by,created_at)
          VALUES(:no,:name,:name_en,:category,:branch,:cc,:supplier,:purchase_date,:cap_date,
          :cost,:residual,:life,:method,0,:nbv,:serial,:location,'نشط',:notes,:uid,:created)""",
          {"no":no,"name":request.form["name"],"name_en":request.form.get("name_en",""),
           "category":category["id"],"branch":request.form.get("branch_id") or None,
           "cc":request.form.get("cost_center_id") or None,
           "supplier":request.form.get("supplier_id") or None,
           "purchase_date":request.form["purchase_date"],
           "cap_date":request.form["capitalization_date"],
           "cost":cost,"residual":residual,
           "life":int(request.form.get("useful_life_months") or category["useful_life_months"]),
           "method":request.form.get("depreciation_method") or category["depreciation_method"],
           "nbv":cost,"serial":request.form.get("serial_number",""),
           "location":request.form.get("location",""),"notes":request.form.get("notes",""),
           "uid":session.get("user_id"),"created":datetime.now()})
        flash(f"تم إنشاء الأصل {no}","success")
        return redirect(url_for("fixed_assets"))
    return render_template("fixed_assets.html",
      assets=rows("""SELECT fa.*,ac.name category_name,b.name branch_name,cc.name cost_center_name
                     FROM fixed_assets fa JOIN asset_categories ac ON ac.id=fa.category_id
                     LEFT JOIN branches b ON b.id=fa.branch_id
                     LEFT JOIN cost_centers cc ON cc.id=fa.cost_center_id
                     ORDER BY fa.id DESC"""),
      categories=rows("SELECT * FROM asset_categories WHERE active=1 ORDER BY code"),
      branches=rows("SELECT id,name FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"),
      suppliers=rows("SELECT id,name FROM suppliers ORDER BY name"))

@app.route("/fixed-assets/depreciation",methods=["GET","POST"])
@login_required
def asset_depreciation():
    if request.method=="POST":
        period_date=request.form["period_date"]
        no=next_asset_run_number(period_date)
        execute("""INSERT INTO asset_depreciation_runs(run_no,period_date,status,notes,created_by,created_at)
                   VALUES(:no,:dt,'مسودة',:notes,:uid,:created)""",
                {"no":no,"dt":period_date,"notes":request.form.get("notes",""),
                 "uid":session.get("user_id"),"created":datetime.now()})
        run_id=row("SELECT id FROM asset_depreciation_runs WHERE run_no=:no",{"no":no})["id"]
        total=0
        for asset in rows("""SELECT * FROM fixed_assets
                             WHERE status='نشط' AND capitalization_date<=:dt""",{"dt":period_date}):
            amount=calculate_monthly_depreciation(asset)
            if amount<=0: continue
            before=float(asset["accumulated_depreciation"] or 0)
            after=round(before+amount,2)
            nbv=round(float(asset["purchase_cost"])-after,2)
            execute("""INSERT INTO asset_depreciation_lines(run_id,asset_id,depreciation_amount,
              accumulated_before,accumulated_after,net_book_value_after)
              VALUES(:run,:asset,:amount,:before,:after,:nbv)""",
              {"run":run_id,"asset":asset["id"],"amount":amount,
               "before":before,"after":after,"nbv":nbv})
            execute("""UPDATE fixed_assets SET accumulated_depreciation=:after,
                       net_book_value=:nbv WHERE id=:id""",
                    {"after":after,"nbv":nbv,"id":asset["id"]})
            total+=amount
        execute("UPDATE asset_depreciation_runs SET total_amount=:total WHERE id=:id",
                {"total":round(total,2),"id":run_id})
        try:
            post_depreciation_run(run_id)
            flash(f"تم تشغيل وترحيل الإهلاك {no}","success")
        except Exception as exc:
            db.session.rollback()
            flash(str(exc),"danger")
        return redirect(url_for("asset_depreciation"))
    return render_template("asset_depreciation.html",
      runs=rows("""SELECT r.*,j.journal_no FROM asset_depreciation_runs r
                   LEFT JOIN journal_entries j ON j.id=r.journal_id
                   ORDER BY r.period_date DESC,r.id DESC"""))

@app.route("/fixed-assets/report")
@login_required
def fixed_assets_report():
    data=rows("""SELECT fa.asset_no,fa.name,ac.name category_name,fa.purchase_date,
      fa.purchase_cost,fa.accumulated_depreciation,fa.net_book_value,
      fa.status,fa.location
      FROM fixed_assets fa JOIN asset_categories ac ON ac.id=fa.category_id
      ORDER BY ac.code,fa.asset_no""")
    return render_template("fixed_assets_report.html",rows=data)

@app.route("/fixed-assets/report.xlsx")
@login_required
def fixed_assets_report_export():
    data=rows("""SELECT fa.asset_no,fa.name,ac.name category_name,fa.purchase_date,
      fa.purchase_cost,fa.accumulated_depreciation,fa.net_book_value,
      fa.status,fa.location
      FROM fixed_assets fa JOIN asset_categories ac ON ac.id=fa.category_id
      ORDER BY ac.code,fa.asset_no""")
    return xlsx_response("fixed_assets.xlsx","الأصول الثابتة",
      ["رقم الأصل","الأصل","الفئة","تاريخ الشراء","التكلفة","مجمع الإهلاك",
       "صافي القيمة الدفترية","الحالة","الموقع"],
      [[r["asset_no"],r["name"],r["category_name"],r["purchase_date"],r["purchase_cost"],
        r["accumulated_depreciation"],r["net_book_value"],r["status"],r["location"]] for r in data])


@app.route("/financial-statements")
@login_required
def financial_statements_center():
    return render_template("financial_statements_center.html")

@app.route("/financial-statements/income-statement")
@login_required
def income_statement():
    filters=financial_date_filters()
    data=financial_statement_data(filters)
    revenues=[x for x in data if x["account_type"]=="إيراد" and abs(x["balance"])>0.005]
    cost_sales=[x for x in data if x["account_type"]=="مصروف" and
                ("تكلفة" in (x["account_name_ar"] or "") or
                 "مبيعات" in (x["account_name_ar"] or "")) and abs(x["balance"])>0.005]
    expenses=[x for x in data if x["account_type"]=="مصروف" and x not in cost_sales
              and abs(x["balance"])>0.005]
    total_revenue=round(sum(x["balance"] for x in revenues),2)
    total_cost=round(sum(x["balance"] for x in cost_sales),2)
    gross_profit=round(total_revenue-total_cost,2)
    total_expenses=round(sum(x["balance"] for x in expenses),2)
    net_profit=round(gross_profit-total_expenses,2)
    return render_template("income_statement.html",filters=filters,revenues=revenues,
      cost_sales=cost_sales,expenses=expenses,total_revenue=total_revenue,
      total_cost=total_cost,gross_profit=gross_profit,total_expenses=total_expenses,
      net_profit=net_profit,branches=rows("SELECT id,name FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"))

@app.route("/financial-statements/balance-sheet")
@login_required
def balance_sheet():
    filters=financial_date_filters()
    data=financial_statement_data(filters)
    assets=[x for x in data if x["account_type"]=="أصل" and abs(x["balance"])>0.005]
    liabilities=[x for x in data if x["account_type"]=="التزام" and abs(x["balance"])>0.005]
    equity=[x for x in data if x["account_type"]=="حقوق ملكية" and abs(x["balance"])>0.005]
    revenue_total=sum(x["balance"] for x in data if x["account_type"]=="إيراد")
    expense_total=sum(x["balance"] for x in data if x["account_type"]=="مصروف")
    retained=round(revenue_total-expense_total,2)
    total_assets=round(sum(x["balance"] for x in assets),2)
    total_liabilities=round(sum(x["balance"] for x in liabilities),2)
    total_equity=round(sum(x["balance"] for x in equity)+retained,2)
    difference=round(total_assets-total_liabilities-total_equity,2)
    return render_template("balance_sheet.html",filters=filters,assets=assets,
      liabilities=liabilities,equity=equity,retained=retained,total_assets=total_assets,
      total_liabilities=total_liabilities,total_equity=total_equity,difference=difference,
      branches=rows("SELECT id,name FROM branches WHERE active=1 ORDER BY name"),
      centers=rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code"))

@app.route("/financial-statements/cash-flow")
@login_required
def cash_flow_statement():
    filters=financial_date_filters()
    conditions=["j.status='مرحّل'"];params={}
    if filters["date_from"]:
        conditions.append("j.journal_date>=:date_from");params["date_from"]=filters["date_from"]
    if filters["date_to"]:
        conditions.append("j.journal_date<=:date_to");params["date_to"]=filters["date_to"]
    data=rows(f"""SELECT a.account_code,a.account_name_ar,
      SUM(l.debit) debit,SUM(l.credit) credit,SUM(l.debit-l.credit) net_movement
      FROM journal_entry_lines l JOIN journal_entries j ON j.id=l.journal_id
      JOIN chart_of_accounts a ON a.id=l.account_id
      WHERE {' AND '.join(conditions)}
      AND a.account_type='أصل'
      AND (a.account_name_ar ILIKE '%صندوق%' OR a.account_name_ar ILIKE '%بنك%'
           OR a.account_name_ar ILIKE '%نقد%')
      GROUP BY a.id,a.account_code,a.account_name_ar ORDER BY a.account_code""",params)
    net_cash=round(sum(float(x["net_movement"] or 0) for x in data),2)
    receipts=row(f"""SELECT COALESCE(SUM(tv.amount),0) amount FROM treasury_vouchers tv
      WHERE tv.voucher_type='قبض' AND tv.posting_status='مرحّل'
      {"AND tv.voucher_date>=:date_from" if filters["date_from"] else ""}
      {"AND tv.voucher_date<=:date_to" if filters["date_to"] else ""}""",params)["amount"]
    payments=row(f"""SELECT COALESCE(SUM(tv.amount),0) amount FROM treasury_vouchers tv
      WHERE tv.voucher_type='صرف' AND tv.posting_status='مرحّل'
      {"AND tv.voucher_date>=:date_from" if filters["date_from"] else ""}
      {"AND tv.voucher_date<=:date_to" if filters["date_to"] else ""}""",params)["amount"]
    return render_template("cash_flow_statement.html",filters=filters,accounts=data,
      receipts=receipts,payments=payments,net_cash=net_cash)

@app.route("/financial-statements/export.xlsx")
@login_required
def financial_statements_export():
    statement=request.args.get("statement","income")
    filters=financial_date_filters()
    data=financial_statement_data(filters)
    if statement=="balance":
        records=[[x["account_code"],x["account_name_ar"],x["account_type"],x["balance"]]
                 for x in data if x["account_type"] in ("أصل","التزام","حقوق ملكية")
                 and abs(x["balance"])>0.005]
        return xlsx_response("balance_sheet.xlsx","الميزانية العمومية",
          ["الكود","الحساب","النوع","الرصيد"],records)
    records=[[x["account_code"],x["account_name_ar"],x["account_type"],x["balance"]]
             for x in data if x["account_type"] in ("إيراد","مصروف")
             and abs(x["balance"])>0.005]
    return xlsx_response("income_statement.xlsx","قائمة الدخل",
      ["الكود","الحساب","النوع","الرصيد"],records)

@app.route("/party-statements")
@login_required
def party_statements():
    party_type=request.args.get("party_type","customer")
    party_id=request.args.get("party_id",type=int)
    date_from=request.args.get("date_from","")
    date_to=request.args.get("date_to","")
    data=party_statement_rows(party_type,party_id,date_from,date_to) if party_id else []
    parties=rows("SELECT id,name FROM customers ORDER BY name") if party_type=="customer" \
            else rows("SELECT id,name FROM suppliers ORDER BY name")
    return render_template("party_statement.html",party_type=party_type,party_id=party_id,
      date_from=date_from,date_to=date_to,rows=data,parties=parties)

@app.route("/party-statements.xlsx")
@login_required
def party_statements_export():
    party_type=request.args.get("party_type","customer")
    party_id=request.args.get("party_id",type=int)
    data=party_statement_rows(party_type,party_id,request.args.get("date_from",""),
                              request.args.get("date_to","")) if party_id else []
    return xlsx_response("party_statement.xlsx","كشف الحساب",
      ["التاريخ","القيد","المرجع","البيان","رقم الفاتورة","مدين","دائن","الرصيد"],
      [[r["journal_date"],r["journal_no"],r["reference"],r["line_description"],
        r["invoice_number"],r["debit"],r["credit"],r["running_balance"]] for r in data])

@app.route("/payables-aging")
@login_required
def payables_aging():
    data=rows("""SELECT s.id,s.name,
      SUM(CASE WHEN CURRENT_DATE-si.invoice_date<=30 THEN si.total ELSE 0 END) bucket_0_30,
      SUM(CASE WHEN CURRENT_DATE-si.invoice_date BETWEEN 31 AND 60 THEN si.total ELSE 0 END) bucket_31_60,
      SUM(CASE WHEN CURRENT_DATE-si.invoice_date BETWEEN 61 AND 90 THEN si.total ELSE 0 END) bucket_61_90,
      SUM(CASE WHEN CURRENT_DATE-si.invoice_date>90 THEN si.total ELSE 0 END) bucket_over_90,
      SUM(si.total) total_outstanding
      FROM supplier_invoices si JOIN suppliers s ON s.id=si.supplier_id
      WHERE si.status='معتمدة'
      GROUP BY s.id,s.name ORDER BY total_outstanding DESC""")
    return render_template("payables_aging.html",rows=data)

@app.route("/executive-dashboard")
@login_required
def executive_dashboard():
    sales=row("""SELECT COALESCE(SUM(total),0) amount,COUNT(*) count
                 FROM invoices WHERE status='معتمدة'""")
    purchases=row("""SELECT COALESCE(SUM(total),0) amount,COUNT(*) count
                     FROM supplier_invoices WHERE status='معتمدة'""")
    expenses=row("""SELECT COALESCE(SUM(amount),0) amount,COUNT(*) count
                    FROM expenses WHERE status='معتمدة'""")
    receivables=row("""SELECT COALESCE(SUM(i.total-COALESCE(
      (SELECT SUM(a.allocated_amount) FROM invoice_payment_allocations a
       WHERE a.invoice_id=i.id),0)),0) amount
      FROM invoices i WHERE i.status='معتمدة'""")
    payables=row("""SELECT COALESCE(SUM(total),0) amount FROM supplier_invoices
                    WHERE status='معتمدة'""")
    inventory_value=row("""SELECT COALESCE(SUM(quantity*cost),0) amount FROM inventory
                           WHERE active=1""")
    cash=row("""SELECT COALESCE(SUM(CASE WHEN l.debit>0 THEN l.debit ELSE -l.credit END),0) amount
      FROM journal_entry_lines l JOIN journal_entries j ON j.id=l.journal_id
      JOIN chart_of_accounts a ON a.id=l.account_id
      WHERE j.status='مرحّل' AND a.account_type='أصل'
      AND (a.account_name_ar ILIKE '%صندوق%' OR a.account_name_ar ILIKE '%بنك%'
           OR a.account_name_ar ILIKE '%نقد%')""")
    top_customers=rows("""SELECT c.name,SUM(i.total) total FROM invoices i
      JOIN customers c ON c.id=i.customer_id WHERE i.status='معتمدة'
      GROUP BY c.id,c.name ORDER BY total DESC LIMIT 10""")
    top_suppliers=rows("""SELECT s.name,SUM(si.total) total FROM supplier_invoices si
      JOIN suppliers s ON s.id=si.supplier_id WHERE si.status='معتمدة'
      GROUP BY s.id,s.name ORDER BY total DESC LIMIT 10""")
    low_stock=rows("""SELECT sku,name,quantity,reorder_level FROM inventory
      WHERE active=1 AND quantity<=reorder_level ORDER BY quantity LIMIT 10""")
    monthly=rows("""SELECT TO_CHAR(invoice_date,'YYYY-MM') month,SUM(total) total
      FROM invoices WHERE status='معتمدة'
      GROUP BY TO_CHAR(invoice_date,'YYYY-MM') ORDER BY month DESC LIMIT 12""")
    net_profit=round(float(sales["amount"] or 0)-float(expenses["amount"] or 0),2)
    return render_template("executive_dashboard.html",sales=sales,purchases=purchases,
      expenses=expenses,receivables=receivables,payables=payables,
      inventory_value=inventory_value,cash=cash,net_profit=net_profit,
      top_customers=top_customers,top_suppliers=top_suppliers,
      low_stock=low_stock,monthly=monthly)


@app.route("/reports")
@login_required
def reports():
    grouped = {}
    for slug, definition in ACCOUNTING_REPORTS.items():
        grouped.setdefault(definition["group"], []).append({"slug": slug, **definition})
    return render_template("reports_center.html", grouped=grouped)

@app.route("/reports/<slug>")
@login_required
def accounting_report(slug):
    if slug not in ACCOUNTING_REPORTS:
        return "التقرير غير موجود", 404
    filters = report_filters()
    context = report_context(slug, filters)
    accounts = rows("SELECT id,account_code,account_name_ar FROM chart_of_accounts WHERE active=1 ORDER BY account_code")
    centers = rows("SELECT id,code,name FROM cost_centers WHERE active=1 ORDER BY code")
    company = row("SELECT * FROM settings WHERE id=1")
    audit("VIEW", "REPORT", f"عرض تقرير {context['definition']['title']}")
    return render_template("accounting_report.html", slug=slug, filters=filters, accounts=accounts,
                           centers=centers, company=company, **context)

@app.route("/reports/<slug>/export.xlsx")
@login_required
def accounting_report_export(slug):
    if slug not in ACCOUNTING_REPORTS:
        return "التقرير غير موجود", 404
    filters = report_filters()
    context = report_context(slug, filters)
    data = context["data"]
    keys_by_slug = {
        "coa-report":["account_code","account_name_ar","account_name_en","account_type","parent_name","level","normal_balance","statement_type","accepts_entries","active"],
        "ledger-summary":["account_code","account_name_ar","total_debit","total_credit","debit_balance","credit_balance","movement_count"],
        "account-balances":["account_code","account_name_ar","total_debit","total_credit","debit_balance","credit_balance","movement_count"],
        "trial-balance":["account_code","account_name_ar","total_debit","total_credit","debit_balance","credit_balance","movement_count"],
        "ledger-detail":["journal_date","journal_no","reference","line_description","invoice_number","invoice_date","debit","credit","running_balance","status"],
        "journal-report":["journal_date","journal_no","status","reference","description","account_code","account_name_ar","debit","credit","party_name","tax_number","invoice_number","invoice_date","cost_center_name"],
        "cost-center-movements-center":["cost_center_name","journal_date","journal_no","account_code","account_name_ar","line_description","debit","credit","running_balance","status"],
        "cc-ledger-detail":["cost_center_name","journal_date","journal_no","account_code","account_name_ar","line_description","debit","credit","running_balance","status"],
        "cost-center-movements-account":["account_code","account_name_ar","cost_center_name","journal_date","journal_no","line_description","debit","credit","status"],
        "cc-ledger-summary":["cost_center_code","cost_center_name","total_debit","total_credit","debit_balance","credit_balance","movement_count"],
        "cc-trial-balance":["cost_center_code","cost_center_name","total_debit","total_credit","debit_balance","credit_balance","movement_count"],
        "cc-balances":["cost_center_code","cost_center_name","total_debit","total_credit","debit_balance","credit_balance","movement_count"],
        "accounts-with-centers":["account_code","account_name_ar","cost_center_code","cost_center_name","total_debit","total_credit","balance"],
        "centers-with-accounts":["cost_center_code","cost_center_name","account_code","account_name_ar","total_debit","total_credit","balance"],
        "vat-report":["journal_date","journal_no","invoice_date","invoice_number","party_name","tax_number","tax_direction","amount_before_tax","vat_amount","total_amount","compliance_status"],
    }
    keys = keys_by_slug[slug]
    records = [[item.get(key) for key in keys] for item in data]
    audit("EXPORT", "REPORT", f"تصدير تقرير {context['definition']['title']}")
    return xlsx_response(f"{slug}.xlsx", context["definition"]["title"][:31], context["headers"], records)


def xlsx_response(filename, sheet_name, headers, records):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True
    ws.append(headers)
    for record in records:
        ws.append(list(record))
    for column in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[column[0].column_letter].width = min(max_length + 3, 45)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


def csv_response(filename, headers, records):
    output = io.StringIO()
    output.write("\ufeff")
    writer = csv.writer(output)
    writer.writerow(headers)
    writer.writerows(records)
    return Response(
        output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.route("/export/invoices.xlsx")
@login_required
def export_invoices():
    data = rows("""
        SELECT i.invoice_no, i.invoice_date, c.name customer_name,
               i.subtotal, i.vat, i.total, i.status
        FROM invoices i
        JOIN customers c ON c.id=i.customer_id
        ORDER BY i.invoice_date DESC, i.id DESC
    """)
    records = [
        [r["invoice_no"], r["invoice_date"], r["customer_name"],
         r["subtotal"], r["vat"], r["total"], r["status"]]
        for r in data
    ]
    return xlsx_response(
        "invoices.xlsx",
        "الفواتير",
        ["رقم الفاتورة", "التاريخ", "العميل", "قبل الضريبة",
         "الضريبة", "الإجمالي", "الحالة"],
        records
    )


@app.route("/export/expenses.xlsx")
@login_required
def export_expenses():
    data = rows("""
        SELECT expense_date, category, description, amount, vat, total
        FROM expenses ORDER BY expense_date DESC, id DESC
    """)
    records = [
        [r["expense_date"], r["category"], r["description"],
         r["amount"], r["vat"], r["total"]]
        for r in data
    ]
    return xlsx_response(
        "expenses.xlsx",
        "المصروفات",
        ["التاريخ", "التصنيف", "الوصف", "قبل الضريبة", "الضريبة", "الإجمالي"],
        records
    )


@app.route("/export/customers.xlsx")
@login_required
def export_customers():
    data = rows("""
        SELECT name, vat_number, phone, email, balance
        FROM customers ORDER BY name
    """)
    records = [
        [r["name"], r["vat_number"], r["phone"], r["email"], r["balance"]]
        for r in data
    ]
    return xlsx_response(
        "customers.xlsx",
        "العملاء",
        ["اسم العميل", "الرقم الضريبي", "الهاتف", "البريد", "الرصيد"],
        records
    )

if __name__ == "__main__":
    with app.app_context():
        init_db()
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
