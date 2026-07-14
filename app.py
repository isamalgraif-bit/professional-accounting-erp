from flask import Flask, render_template, request, redirect, url_for, session, flash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from functools import wraps
from datetime import datetime
import os

app = Flask(__name__, template_folder=".")
app.secret_key = os.environ.get("SECRET_KEY", "development-only-change-me")

database_url = os.environ.get("DATABASE_URL", "sqlite:///erp.db")
if database_url.startswith("postgres://"):
    database_url = database_url.replace(
        "postgres://",
        "postgresql+psycopg://",
        1
    )
elif database_url.startswith("postgresql://"):
    database_url = database_url.replace(
        "postgresql://",
        "postgresql+psycopg://",
        1
    )
    

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
            currency VARCHAR(10) DEFAULT 'SAR'
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
    if request.method == "POST":
        execute("""UPDATE settings SET company_name_ar=:ar, company_name_en=:en, vat_number=:vat,
                   cr_number=:cr, currency=:cur WHERE id=1""",
                {"ar": request.form["company_name_ar"], "en": request.form["company_name_en"],
                 "vat": request.form["vat_number"], "cr": request.form["cr_number"],
                 "cur": request.form["currency"]})
        flash("تم حفظ إعدادات الشركة", "success")
        return redirect(url_for("settings"))
    return render_template("settings.html", row=row("SELECT * FROM settings WHERE id=1"))

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
        subtotal = float(request.form["subtotal"] or 0)
        vat = round(subtotal * 0.15, 2)
        total = round(subtotal + vat, 2)
        invoice_no = f"INV-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
        execute("""INSERT INTO invoices(invoice_no,customer_id,invoice_date,subtotal,vat,total,
                   status,branch_id,notes,created_at)
                   VALUES(:no,:cid,:dt,:sub,:vat,:tot,:st,:bid,:notes,:created)""",
                {"no": invoice_no, "cid": request.form["customer_id"], "dt": request.form["invoice_date"],
                 "sub": subtotal, "vat": vat, "tot": total, "st": request.form["status"],
                 "bid": request.form["branch_id"], "notes": request.form["notes"],
                 "created": datetime.now()})
        flash("تم إنشاء الفاتورة واحتساب الضريبة 15٪", "success")
    invoice_rows = rows("""SELECT i.*, c.name customer_name, b.name branch_name
                           FROM invoices i JOIN customers c ON c.id=i.customer_id
                           LEFT JOIN branches b ON b.id=i.branch_id ORDER BY i.id DESC""")
    return render_template("invoices.html", rows=invoice_rows,
                           customers=rows("SELECT * FROM customers ORDER BY name"),
                           branches=rows("SELECT * FROM branches WHERE active=1 ORDER BY name"))

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

if __name__ == "__main__":
    with app.app_context():
        init_db()
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
