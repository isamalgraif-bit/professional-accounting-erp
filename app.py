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

APP_VERSION = "1.4.0"

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
            vat_number VARCHAR(50),
            phone VARCHAR(50),
            email VARCHAR(255),
            balance NUMERIC(18,2) DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS suppliers(
            id SERIAL PRIMARY KEY,
            name VARCHAR(255) NOT NULL,
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
        "ALTER TABLE journal_entry_lines ADD COLUMN IF NOT EXISTS tax_number VARCHAR(50) DEFAULT ''"
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
    db.session.commit()

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

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
                       logo_mime=:logo_mime, vat_rate=:vat_rate
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
                 "vat_rate": float(request.form.get("vat_rate", 15) or 15)})

        audit("UPDATE", "SETTINGS", "تم تحديث إعدادات الشركة والشعار")
        flash("تم حفظ إعدادات الشركة والشعار", "success")
        return redirect(url_for("settings"))

    return render_template("settings.html", row=current)

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
        execute("""INSERT INTO customers(name,vat_number,phone,email)
                   VALUES(:n,:v,:p,:e)""",
                {"n": request.form["name"], "v": request.form["vat_number"],
                 "p": request.form["phone"], "e": request.form["email"]})
        audit("CREATE", "CUSTOMER", f"إضافة عميل: {request.form['name']}")
        flash("تمت إضافة العميل", "success")
    return render_template("customers.html", rows=rows("SELECT * FROM customers ORDER BY id DESC"))

@app.route("/suppliers", methods=["GET","POST"])
@login_required
def suppliers():
    if request.method == "POST":
        execute("""INSERT INTO suppliers(name,vat_number,phone,email)
                   VALUES(:n,:v,:p,:e)""",
                {"n": request.form["name"], "v": request.form["vat_number"],
                 "p": request.form["phone"], "e": request.form["email"]})
        audit("CREATE", "SUPPLIER", f"إضافة مورد: {request.form['name']}")
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

        audit("CREATE", "INVOICE", f"إنشاء فاتورة: {invoice_no}")
        flash("تم إنشاء الفاتورة بجميع البنود بنجاح", "success")
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
    if request.method == "POST":
        amount = float(request.form["amount"] or 0)
        vat = round(amount * 0.15, 2) if request.form.get("taxable") == "1" else 0
        total = round(amount + vat, 2)
        execute("""INSERT INTO expenses(expense_date,category,description,amount,vat,total,branch_id)
                   VALUES(:dt,:cat,:des,:amt,:vat,:tot,:bid)""",
                {"dt": request.form["expense_date"], "cat": request.form["category"],
                 "des": request.form["description"], "amt": amount, "vat": vat,
                 "tot": total, "bid": request.form["branch_id"]})
        audit("CREATE", "EXPENSE", f"إضافة مصروف: {request.form['category']}")
        flash("تم تسجيل المصروف", "success")
    expense_rows = rows("""SELECT e.*, b.name branch_name FROM expenses e
                           LEFT JOIN branches b ON b.id=e.branch_id ORDER BY e.id DESC""")
    return render_template("expenses.html", rows=expense_rows,
                           branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"))

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
        SELECT i.*, c.name customer_name, c.vat_number customer_vat,
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
        qr_payload=tlv
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
            parent_id,level,accepts_entries,active)
            VALUES(:code,:ar,:en,:type,:parent,:level,:accepts,:active)""",
            {"code":request.form["account_code"].strip(),
             "ar":request.form["account_name_ar"].strip(),
             "en":request.form.get("account_name_en","").strip(),
             "type":request.form["account_type"],
             "parent":parent_id,
             "level":parent_level+1,
             "accepts":1 if request.form.get("accepts_entries")=="1" else 0,
             "active":1 if request.form.get("active")=="1" else 0})
        flash("تمت إضافة الحساب","success")
        return redirect(url_for("chart_of_accounts"))
    accounts = rows("""SELECT a.*,p.account_name_ar parent_name
                       FROM chart_of_accounts a
                       LEFT JOIN chart_of_accounts p ON p.id=a.parent_id
                       ORDER BY a.account_code""")
    return render_template("chart_of_accounts.html",accounts=accounts)

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
    return render_template("multi_journal_view.html", batch=batch, groups=groups, lines=lines, company=company)

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



@app.route("/reports")
@login_required
def reports():
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    invoice_where = []
    expense_where = []
    params = {}

    if date_from:
        invoice_where.append("invoice_date >= :date_from")
        expense_where.append("expense_date >= :date_from")
        params["date_from"] = date_from
    if date_to:
        invoice_where.append("invoice_date <= :date_to")
        expense_where.append("expense_date <= :date_to")
        params["date_to"] = date_to

    invoice_filter = (" WHERE " + " AND ".join(invoice_where)) if invoice_where else ""
    expense_filter = (" WHERE " + " AND ".join(expense_where)) if expense_where else ""

    sales_data = row(f"""
        SELECT COALESCE(SUM(subtotal),0) subtotal,
               COALESCE(SUM(vat),0) vat,
               COALESCE(SUM(total),0) total,
               COUNT(*) count
        FROM invoices {invoice_filter}
    """, params)

    expense_data = row(f"""
        SELECT COALESCE(SUM(amount),0) amount,
               COALESCE(SUM(vat),0) vat,
               COALESCE(SUM(total),0) total,
               COUNT(*) count
        FROM expenses {expense_filter}
    """, params)

    inventory_value = row("""
        SELECT COALESCE(SUM(quantity * cost),0) value,
               COUNT(*) count
        FROM inventory
    """)
    payroll = row("""
        SELECT COALESCE(SUM(basic_salary + allowances),0) total,
               COUNT(*) count
        FROM employees WHERE active=1
    """)
    top_customers = rows(f"""
        SELECT c.name, COUNT(i.id) invoice_count,
               COALESCE(SUM(i.total),0) total
        FROM customers c
        LEFT JOIN invoices i ON i.customer_id=c.id
        {"AND " + " AND ".join("i." + x for x in invoice_where) if invoice_where else ""}
        GROUP BY c.id, c.name
        ORDER BY total DESC
        LIMIT 10
    """, params)

    net_result = float(sales_data["subtotal"]) - float(expense_data["amount"])
    net_vat = float(sales_data["vat"]) - float(expense_data["vat"])

    return render_template(
        "reports.html",
        sales=sales_data,
        expenses=expense_data,
        inventory_value=inventory_value,
        payroll=payroll,
        net_result=net_result,
        net_vat=net_vat,
        top_customers=top_customers,
        date_from=date_from,
        date_to=date_to,
    )



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
