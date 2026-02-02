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
app.config['SECRET_KEY'] = 'ZEN_AI_STRUCTURA_FINAL_V14_SMART_LAYOUT' 
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = datetime.timedelta(days=30)

# ==============================================================================
# 2. DATABASE MODELS
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
    return float(matches[0]) if matches else 0.0

# ==============================================================================
# 4. MATH ENGINE (HANDLES BOTH MODES)
# ==============================================================================
def recalculate_math(data):
    items = data.get("items", [])
    footer = data.get("footer", {})
    # LOGIC FIX: Check if layout is explicitly personal
    layout = data.get("layout", "business")
    is_personal = (layout == "personal")
    running_total = 0.0
    
    global_tax_pct = 0.0
    if not is_personal:
        tax_str = str(footer.get("tax_summary") or "")
        rates = re.findall(r"(\d+(?:\.\d+)?)", tax_str)
        if rates:
            raw = [float(r) for r in rates if float(r) <= 50]
            if raw:
                s = sum(raw)
                global_tax_pct = s if any(abs(s - x) < 0.1 for x in [5,12,18,28]) else max(raw)
        if abs(global_tax_pct - 9.0) < 0.1: global_tax_pct = 18.0

    for item in items:
        qty = extract_number(item.get("quantity")) or 1.0
        rate = extract_number(item.get("rate"))
        disc = extract_number(item.get("discount_percent"))
        desc = str(item.get("particulars") or "").lower()

        if ("discount" in desc or "less" in desc) and rate > 0: rate *= -1
        
        # MATH: Personal mode doesn't do reverse tax calculation
        gross = qty * rate
        taxable = gross * (1 - (disc / 100.0))
        
        if is_personal:
            tax_amt = 0
            final_amt = taxable
            item["tax_rate"] = "0%"
        else:
            tax_amt = taxable * (global_tax_pct / 100.0)
            final_amt = taxable + tax_amt
            item["tax_rate"] = f"{int(global_tax_pct)}%"

        item.update({"quantity": qty, "rate": rate, "amount": round(final_amt, 2)})
        running_total += final_amt

    footer["total_amount"] = round(running_total, 2)
    return data

# ==============================================================================
# 5. AI PARSING LOGIC (THE SMART SWITCH)
# ==============================================================================
def parse_invoice_vision(image_path, user_instruction=""):
    base64_img = encode_image(image_path)
    
    # THE FIX: This prompt logic decides layout based on input type
    prompt = f"""
    Extract data into JSON.
    RULES:
    1. IF input is a Formal Tax Invoice -> Set "layout": "business", "company_name": "SHARMA ENTERPRISES".
    2. IF input is Informal Text (e.g. "{user_instruction}") OR a handwritten list -> Set "layout": "personal", "company_name": "EXPENSE SHEET".
    
    USER INSTRUCTION: '{user_instruction}'
    """
    
    payload = {"model": MODEL, "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}]}]}
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    
    r = requests.post(API_URL, headers=headers, json=payload, timeout=60)
    raw = r.json()["choices"][0]["message"]["content"]
    json_str = re.sub(r'[\x00-\x09\x0b-\x1f\x7f]', '', raw[raw.find("{"):raw.rfind("}")+1])
    return recalculate_math(json.loads(json_str, strict=False))

# ==============================================================================
# 6. EXCEL LAYOUTS (PERSONAL VS BUSINESS)
# ==============================================================================
def write_personal_layout(ws, data):
    items, foot = data.get("items", []), data.get("footer", {})
    # SIMPLE HEADER (Matches Image 2)
    ws['A1'] = "EXPENSE SHEET"
    ws['A1'].font = Font(size=16, bold=True, color="000000")
    
    # Simple Columns: Description, Qty, Rate, Amount
    headers = ["Description", "Quantity", "Rate", "Amount"]
    widths = [40, 10, 15, 15]
    
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(bold=True)
        ws.column_dimensions[get_column_letter(i)].width = w

    curr = 5
    for item in items:
        ws.cell(row=curr, column=1, value=clean(item.get("particulars")))
        ws.cell(row=curr, column=2, value=item.get("quantity"))
        ws.cell(row=curr, column=3, value=item.get("rate"))
        ws.cell(row=curr, column=4, value=item.get("amount"))
        curr += 1
    
    ws.cell(row=curr+1, column=3, value="TOTAL").font = Font(bold=True)
    ws.cell(row=curr+1, column=4, value=foot.get("total_amount")).font = Font(bold=True)

def write_business_layout(ws, data):
    head, items, foot = data.get("header", {}), data.get("items", []), data.get("footer", {})
    last_col = "H"
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    box_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

    ws.merge_cells(f'A1:{last_col}1')
    ws['A1'].value = clean(head.get("company_name")) or "SHARMA ENTERPRISES"
    ws['A1'].font = Font(size=22, bold=True); ws['A1'].alignment = center
    ws['A1'].fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")

    ws.merge_cells(f'A2:{last_col}2')
    ws['A2'].value = clean(head.get("company_subtext")) or "Tax Invoice / GST Extraction"
    ws['A2'].alignment = center

    headers = ["S.N.", "Particulars", "HSN/SAC", "Qty", "Rate", "Gross", "Tax %", "Total"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=6, column=i, value=h)
        c.font = Font(bold=True); c.border = box_border; ws.column_dimensions[get_column_letter(i)].width = 15

    curr = 7
    for item in items:
        vals = [item.get("sn"), item.get("particulars"), item.get("hsn_sac"), item.get("quantity"), item.get("rate"), item.get("gross_amount"), item.get("tax_rate"), item.get("amount")]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=curr, column=i, value=clean(v)); c.border = box_border
        curr += 1

    ws.merge_cells(f'A{curr}:G{curr}')
    ws.cell(row=curr, column=1, value="Total Amount (Inc. GST)").alignment = Alignment(horizontal='right')
    ws.cell(row=curr, column=8, value=clean(foot.get("total_amount"))).font = Font(bold=True); ws.cell(row=curr, column=8).border = box_border

    curr += 1
    ws.merge_cells(f'A{curr}:{last_col}{curr}')
    ws.cell(row=curr, column=1, value=f"Amount in Words: {clean(foot.get('amount_in_words'))}").font = Font(italic=True)

# ==============================================================================
# 7. ROUTES & RENDER BINDING
# ==============================================================================
@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response

@app.route("/")
def home(): return render_template("index.html", user=current_user)

@app.route("/login", methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('home'))
    if request.method == 'POST':
        u = User.query.filter_by(email=request.form.get('email')).first()
        if u and check_password_hash(u.password, request.form.get('password')):
            login_user(u); return redirect(url_for('input_page'))
    return render_template('login.html')

@app.route("/signup", methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        new_u = User(name=request.form.get('name'), email=request.form.get('email'), gender=request.form.get('gender'),
                     password=generate_password_hash(request.form.get('password'), method='scrypt'))
        db.session.add(new_u); db.session.commit()
        return redirect(url_for('login'))
    return render_template('signup.html')

@app.route("/logout")
@login_required
def logout(): logout_user(); return redirect(url_for('home'))

@app.route("/input")
def input_page(): 
    usage = session.get('usage_count', 0)
    if current_user.is_authenticated:
        return render_template("input.html", user=current_user, trials_left=None)
    return render_template("input.html", user=None, trials_left=3-usage)

@app.route("/process", methods=["POST"])
def process():
    if not current_user.is_authenticated and session.get('usage_count', 0) >= 3:
        return jsonify({"error": "3 trials ended, please log in"}), 403
    
    file = request.files.get("image")
    prompt_text = request.form.get("prompt", "")
    
    # Create temp image for prompt-only requests (Zomato fix)
    img_path = os.path.join(UPLOAD_DIR, secure_filename(file.filename) if file else "prompt_bg.png")
    if file: file.save(img_path)
    else: Image.new('RGB', (100, 100), color='white').save(img_path)

    try:
        data = parse_invoice_vision(img_path, prompt_text)
        wb = Workbook(); ws = wb.active
        
        # SMART LAYOUT SWITCH
        if data.get("layout") == "personal":
            write_personal_layout(ws, data)
        else:
            write_business_layout(ws, data)
            
        save_path = os.path.join(UPLOAD_DIR, "Structura_Data.xlsx"); wb.save(save_path)
        
        if current_user.is_authenticated:
            new_h = History(user_id=current_user.id, filename="Structura_Data.xlsx", prompt=prompt_text)
            db.session.add(new_h); db.session.commit()
        else: session['usage_count'] = session.get('usage_count', 0) + 1
        
        return jsonify({"status": "ok", "filename": "Structura_Data.xlsx", "trials_left": 3-session.get('usage_count', 0)})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route("/download")
def download():
    path = os.path.join(UPLOAD_DIR, secure_filename(request.args.get('filename', "Structura_Data.xlsx")))
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
