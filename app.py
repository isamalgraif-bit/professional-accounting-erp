from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from flask_sqlalchemy import SQLAlchemy
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
        "ALTER TABLE invoices ADD COLUMN IF NOT EXISTS invoice_uuid VARCHAR(64)"
    ]
    for migration in migrations:
        db.session.execute(text(migration))

    db.session.execute(text("""
        INSERT INTO users(id, username, password, role)
        VALUES(1,'admin','admin123','admin')
        ON CONFLICT (id) DO NOTHING
    """))
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
    return {"app_settings": settings}

@app.route("/health")
def health():
    db.session.execute(text("SELECT 1"))
    return {"status": "ok"}, 200

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()
        user = row("SELECT * FROM users WHERE username=:u AND password=:p",
                   {"u": username, "p": password})
        if user:
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            return redirect(url_for("dashboard"))
        flash("بيانات الدخول غير صحيحة", "danger")
    return render_template("login.html")

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

        flash("تم حفظ إعدادات الشركة والشعار", "success")
        return redirect(url_for("settings"))

    return render_template("settings.html", row=current)

@app.route("/branches", methods=["GET","POST"])
@login_required
def branches():
    if request.method == "POST":
        execute("INSERT INTO branches(name,city) VALUES(:n,:c)",
                {"n": request.form["name"], "c": request.form["city"]})
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

        invoice_no = f"INV-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
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
        flash("تمت إضافة الموظف", "success")
    employee_rows = rows("""SELECT e.*, b.name branch_name FROM employees e
                            LEFT JOIN branches b ON b.id=e.branch_id ORDER BY e.id DESC""")
    return render_template("employees.html", rows=employee_rows,
                           branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"))



def zatca_tlv_base64(seller_name, vat_number, invoice_datetime, total, vat_total):
    """Create Base64 TLV data for the five standard invoice QR fields."""
    values = [
        (1, str(seller_name or "")),
        (2, str(vat_number or "")),
        (3, str(invoice_datetime or "")),
        (4, f"{Decimal(str(total)):.2f}"),
        (5, f"{Decimal(str(vat_total)):.2f}")
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
