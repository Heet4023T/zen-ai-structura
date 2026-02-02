import os
import json
import requests
import base64
import re
import datetime
from flask import Flask, request, render_template, jsonify, send_file, redirect, url_for, flash, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

# ==============================================================================
# 1. CONFIGURATION & SETUP
# ==============================================================================

API_KEY = os.environ.get("GITHUB_TOKEN")
MODEL = "gpt-4o" 
API_URL = "https://models.inference.ai.azure.com/chat/completions"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
LAST_GENERATED_FILE = None

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config['SECRET_KEY'] = 'FINAL_FULL_CODE_RESTORED_V13_RENDER' 
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = datetime.timedelta(days=30)

# ==============================================================================
# 2. DATABASE & AUTHENTICATION MODELS
# ==============================================================================

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150))
    email = db.Column(db.String(150), unique=True)
    password = db.Column(db.String(150))
    gender = db.Column(db.String(50))

class History(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    filename = db.Column(db.String(300))
    prompt = db.Column(db.Text)
    timestamp = db.Column(db.DateTime, default=datetime.datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

with app.app_context():
    db.create_all()

# ==============================================================================
# 3. UTILITY FUNCTIONS
# ==============================================================================

def clean(v):
    if v is None: return ""
    s = str(v).strip()
    if s.lower() in ["null", "none", "n/a", "", "[]", "{}"]: return ""
    return s

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def extract_number(value):
    if not value: return 0.0
    matches = re.findall(r"(-?\d+(?:\.\d+)?)", str(value).replace(",", ""))
    if matches:
        try: return float(matches[0])
        except ValueError: return 0.0
    return 0.0

def generate_error_excel(error_msg, save_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Error Log"
    ws['A1'] = "PROCESSING FAILED"
    ws['A1'].font = Font(color="FF0000", size=14, bold=True)
    ws['A2'] = str(error_msg)
    ws.column_dimensions['A'].width = 60
    wb.save(save_path)

# ==============================================================================
# 4. MATH ENGINE (SHARMA, BLUEORBIT, JAN AUSHADI LOGIC)
# ==============================================================================

def recalculate_math(data):
    items = data.get("items", [])
    footer = data.get("footer", {})
    layout = data.get("layout", "business")
    is_personal_mode = (layout == "personal")
    running_total = 0.0
    STANDARD_RATES = [0.0, 5.0, 12.0, 18.0, 28.0]

    global_tax_pct = 0.0
    if not is_personal_mode:
        tax_summary = str(footer.get("tax_summary") or "")
        global_rates = re.findall(r"(\d+(?:\.\d+)?)", tax_summary)
        if global_rates:
            raw_nums = [float(r) for r in global_rates if float(r) <= 50]
            if raw_nums:
                s, m = sum(raw_nums), max(raw_nums)
                global_tax_pct = s if any(abs(s - x) < 0.1 for x in STANDARD_RATES) else m
        
        if abs(global_tax_pct - 9.0) < 0.1: global_tax_pct = 18.0
        elif abs(global_tax_pct - 6.0) < 0.1: global_tax_pct = 12.0

    for item in items:
        qty = extract_number(item.get("quantity"))
        rate = extract_number(item.get("rate"))
        disc_pct = extract_number(item.get("discount_percent"))
        desc = str(item.get("particulars") or "").lower()

        if ("discount" in desc or "adjustment" in desc or "less" in desc) and rate > 0:
            rate = -1 * abs(rate)
        if qty == 0 and rate != 0: qty = 1.0

        if rate != 0: gross_amount = qty * rate
        else:
            gross_amount = extract_number(item.get("amount"))
            if ("discount" in desc or "adjustment" in desc) and gross_amount > 0:
                gross_amount = -1 * abs(gross_amount)

        discount_amount = gross_amount * (disc_pct / 100.0) if disc_pct > 0 else 0.0
        taxable_value = gross_amount - discount_amount
        applicable_tax_pct = 0.0
        
        if not is_personal_mode:
            item_tax_str = str(item.get("tax_rate") or "")
            item_tax_nums = re.findall(r"(\d+(?:\.\d+)?)", item_tax_str)
            if item_tax_nums:
                nums = [float(r) for r in item_tax_nums if float(r) <= 100]
                if nums:
                    s, m = sum(nums), max(nums)
                    applicable_tax_pct = s if any(abs(s - x) < 0.1 for x in STANDARD_RATES) else m
            
            if applicable_tax_pct == 0 and global_tax_pct > 0: applicable_tax_pct = global_tax_pct
            if abs(applicable_tax_pct - 9.0) < 0.1: applicable_tax_pct = 18.0

        calc_factor = applicable_tax_pct if applicable_tax_pct < 1.0 else applicable_tax_pct / 100.0
        display_pct = applicable_tax_pct if applicable_tax_pct >= 1.0 else applicable_tax_pct * 100
        tax_amount_val = taxable_value * calc_factor
        final_item_total = taxable_value + tax_amount_val

        item.update({
            "quantity": qty, 
            "rate": rate, 
            "gross_amount": round(gross_amount, 2), 
            "discount_amount": round(discount_amount, 2), 
            "amount": round(final_item_total, 2)
        })
        
        if is_personal_mode: item["amount"] = round(taxable_value, 2)
        else: item["tax_rate"] = f"{int(display_pct)}%" if display_pct > 0 else "0%"
        running_total += item["amount"]

    footer["total_amount"] = round(running_total, 2)
    return data

# ==============================================================================
# 5. LLM / AI PARSING LOGIC
# ==============================================================================

def parse_invoice_vision(image_path, user_instruction=""):
    base64_image = encode_image(image_path)
    prompt = f"Extract data into JSON structure. USER INSTRUCTION: '{user_instruction}'"
    payload = {"model": MODEL, "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}]}
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    
    try:
        r = requests.post(API_URL, headers=headers, json=payload, timeout=60)
        raw = r.json()["choices"][0]["message"]["content"]
        json_str = re.sub(r'[\x00-\x09\x0b-\x1f\x7f]', '', raw[raw.find("{"):raw.rfind("}")+1])
        return recalculate_math(json.loads(json_str, strict=False))
    except Exception as e: raise e 

# ==============================================================================
# 6. EXCEL LAYOUTS
# ==============================================================================

def write_business_layout(ws, data):
    head, items, foot = data.get("header", {}), data.get("items", []), data.get("footer", {})
    has_disc = any(item.get("discount_amount", 0) > 0 for item in items)
    
    headers = ["S.N.", "Particulars", "HSN/SAC", "Qty", "Rate"]
    keys = ["sn", "particulars", "hsn_sac", "quantity", "rate"]
    if has_disc:
        headers.extend(["Gross Amt", "Discount"])
        keys.extend(["gross_amount", "discount_amount"])
    headers.extend(["Tax %", "Amount (Inc. Tax)"])
    keys.extend(["tax_rate", "amount"])

    num_cols = len(headers)
    last_col = get_column_letter(num_cols)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    box_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

    ws.merge_cells(f'A1:{last_col}1')
    ws['A1'].value = clean(head.get("company_name")) or "INVOICE"
    ws['A1'].font = Font(size=22, bold=True)
    ws['A1'].alignment = center
    ws['A1'].fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")

    for i, h in enumerate(headers, 1):
        c = ws.cell(row=6, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = center
        c.border = box_border
        ws.column_dimensions[get_column_letter(i)].width = 20

    curr = 7
    for item in items:
        for i, k in enumerate(keys, 1):
            c = ws.cell(row=curr, column=i, value=clean(item.get(k)))
            c.border = box_border
        curr += 1

    ws.merge_cells(f'A{curr}:{get_column_letter(num_cols-1)}{curr}')
    ws.cell(row=curr, column=1, value="Total Amount (Inc. GST)").alignment = Alignment(horizontal='right')
    ws.cell(row=curr, column=num_cols, value=clean(foot.get("total_amount"))).font = Font(bold=True)

def write_personal_layout(ws, data):
    items, foot = data.get("items", []), data.get("footer", {})
    ws['A1'] = "EXPENSE SHEET"
    ws['A1'].font = Font(size=16, bold=True)
    headers = ["Description", "Quantity", "Rate", "Amount"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h).font = Font(bold=True)
    
    curr = 5
    for item in items:
        ws.cell(row=curr, column=1, value=clean(item.get("particulars")))
        ws.cell(row=curr, column=2, value=clean(item.get("quantity")))
        ws.cell(row=curr, column=3, value=clean(item.get("rate")))
        ws.cell(row=curr, column=4, value=clean(item.get("amount")))
        curr += 1
    
    ws.cell(row=curr+1, column=3, value="TOTAL").font = Font(bold=True)
    ws.cell(row=curr+1, column=4, value=clean(foot.get("total_amount"))).font = Font(bold=True)

# ==============================================================================
# 7. ROUTES & RENDER BINDING
# ==============================================================================

@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response

def generate_excel(data, save_path):
    wb = Workbook()
    ws = wb.active
    if data.get("layout") == "personal": write_personal_layout(ws, data)
    else: write_business_layout(ws, data)
    wb.save(save_path)

@app.route("/")
def home(): return render_template("index.html", user=current_user)

@app.route("/login", methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('home'))
    if request.method == 'POST':
        u = User.query.filter_by(email=request.form.get('email')).first()
        if u and check_password_hash(u.password, request.form.get('password')):
            login_user(u)
            return redirect(url_for('input_page'))
    return render_template('login.html')

@app.route("/signup", methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        new_u = User(name=request.form.get('name'), email=request.form.get('email'), gender=request.form.get('gender'),
                     password=generate_password_hash(request.form.get('password'), method='scrypt'))
        db.session.add(new_u)
        db.session.commit()
        return redirect(url_for('login'))
    return render_template('signup.html')

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for('home'))

@app.route("/input")
def input_page(): 
    # 1. Get the current usage (default to 0 if new visitor)
    usage = session.get('usage_count', 0)
    
    # 2. Check if user is actually logged in
    # This is the "Safety Check" that prevents the crash
    if current_user.is_authenticated:
        # User is logged in, send the actual user object
        return render_template("input.html", user=current_user, trials_left=None)
    else:
        # No one is logged in, send 'None' for user
        # This tells the HTML to show "Log In" instead of the Avatar
        return render_template("input.html", user=None, trials_left=3-usage)

@app.route("/process", methods=["POST"])
def process():
    if not current_user.is_authenticated and session.get('usage_count', 0) >= 3:
        return jsonify({"error": "3 trials ended now do log in"}), 403
    file = request.files.get("image")
    img_path = os.path.join(UPLOAD_DIR, secure_filename(file.filename))
    file.save(img_path)
    try:
        data = parse_invoice_vision(img_path, request.form.get("prompt", ""))
        save_path = os.path.join(UPLOAD_DIR, f"{file.filename}.xlsx")
        generate_excel(data, save_path)
        global LAST_GENERATED_FILE
        LAST_GENERATED_FILE = save_path
        if not current_user.is_authenticated: session['usage_count'] = session.get('usage_count', 0) + 1
        return jsonify({"status": "ok", "filename": f"{file.filename}.xlsx", "trials_left": 3-session.get('usage_count', 0)})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route("/download")
def download():
    path = os.path.join(UPLOAD_DIR, secure_filename(request.args.get('filename')))
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

