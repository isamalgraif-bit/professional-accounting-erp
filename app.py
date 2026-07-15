from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import text
from functools import wraps
from datetime import datetime
import os
import csv
import io
import base64
import uuid
from decimal import Decimal
import qrcode
from openpyxl import Workbook

APP_VERSION = "5.0.2"

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
    if request.method == "POST":
        execute("""INSERT INTO inventory(sku,name,quantity,unit,cost,sale_price,reorder_level)
                   VALUES(:sku,:n,:q,:u,:c,:sp,:rl)""",
                {"sku": request.form["sku"] or None, "n": request.form["name"],
                 "q": float(request.form["quantity"] or 0), "u": request.form["unit"],
                 "c": float(request.form["cost"] or 0), "sp": float(request.form["sale_price"] or 0),
                 "rl": float(request.form["reorder_level"] or 0)})
        audit("CREATE", "INVENTORY", f"إضافة صنف: {request.form['name']}")
        flash("تمت إضافة الصنف", "success")
    return render_template("inventory.html", rows=rows("SELECT * FROM inventory ORDER BY id DESC"))

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
