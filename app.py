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
app.config['SECRET_KEY'] = 'FINAL_SHARMA_LAYOUT_FORCED_V17' 
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

def generate_error_excel(error_msg, save_path):
    wb = Workbook()
    ws = wb.active
    ws['A1'] = "PROCESSING FAILED"
    ws['A1'].font = Font(color="FF0000", size=14, bold=True)
    ws['A2'] = str(error_msg)
    ws.column_dimensions['A'].width = 60
    wb.save(save_path)

# ==============================================================================
# 4. MATH ENGINE
# ==============================================================================
def recalculate_math(data):
    items = data.get("items", [])
    footer = data.get("footer", {})
    is_personal_mode = (data.get("layout") == "personal")
    running_total = 0.0
    
    global_tax_pct = 0.0
    if not is_personal_mode:
        tax_str = str(footer.get("tax_summary") or "")
        rates = re.findall(r"(\d+(?:\.\d+)?)", tax_str)
        if rates:
            raw = [float(r) for r in rates if float(r) <= 50]
            if raw:
                s = sum(raw)
                global_tax_pct = s if any(abs(s - x) < 0.1 for x in [5,12,18,28]) else max(raw)
        if abs(global_tax_pct - 9.0) < 0.1: global_tax_pct = 18.0

    for item in items:
        qty = extract_number(item.get("quantity"))
        rate = extract_number(item.get("rate"))
        disc = extract_number(item.get("discount_percent"))
        desc = str(item.get("particulars") or "").lower()

        if ("discount" in desc or "less" in desc) and rate > 0: rate *= -1
        if qty == 0 and rate != 0: qty = 1.0

        gross = qty * rate
        # Force tax calculation on every row if in business mode
        if not is_personal_mode:
            taxable = gross * (1 - (disc / 100.0))
            tax_amt = taxable * (global_tax_pct / 100.0)
            final_amt = taxable + tax_amt
            item["tax_rate"] = f"{int(global_tax_pct)}%"
        else:
            final_amt = gross
            item["tax_rate"] = "0%"

        item.update({"quantity": qty, "rate": rate, "amount": round(final_amt, 2)})
        running_total += final_amt

    footer["total_amount"] = round(running_total, 2)
    return data

# ==============================================================================
# 5. AI PARSING LOGIC
# ==============================================================================
def parse_invoice_vision(image_path, user_instruction=""):
    base64_img = encode_image(image_path)
    
    prompt = f"""
    Extract invoice data. USER INSTRUCTION: "{user_instruction}"
    
    RULES:
    1. **PERSONAL**: If input is text like "paid 500 to zomato", set "layout": "personal".
    2. **BUSINESS**: If input is an image of a bill, set "layout": "business".
       - **COMPANY NAME**: Extract the big title at the top (e.g. "Sharma Enterprises").
       - **SUBTEXT**: Extract address/GSTIN under the title.
       - **BANK**: Extract Bank Name, A/c No, IFSC.
    
    JSON STRUCTURE:
    {{
      "layout": "business",
      "header": {{ "company_name": null, "company_subtext": null, "gstin": null, "buyer_name": null, "invoice_no": null, "date": null, "bank_details": {{ "bank_name": null, "acc_no": null, "ifsc": null }} }},
      "items": [ {{ "sn": 1, "particulars": null, "hsn_sac": null, "quantity": 0, "rate": 0, "amount": 0 }} ],
      "footer": {{ "tax_summary": "18%", "total_amount": 0, "amount_in_words": null }}
    }}
    """
    
    payload = {"model": MODEL, "messages": [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}}]}]}
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    
    r = requests.post(API_URL, headers=headers, json=payload, timeout=60)
    raw = r.json()["choices"][0]["message"]["content"]
    json_str = re.sub(r'[\x00-\x09\x0b-\x1f\x7f]', '', raw[raw.find("{"):raw.rfind("}")+1])
    return recalculate_math(json.loads(json_str, strict=False))

# ==============================================================================
# 6. EXCEL LAYOUTS (STRICTLY ENFORCED FORMAT)
# ==============================================================================
def write_business_layout(ws, data):
    head, items, foot = data.get("header", {}), data.get("items", []), data.get("footer", {})
    
    # SETUP STYLES
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    box_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))
    
    # 1. ORANGE HEADER (A1:H1) - FORCED
    ws.merge_cells('A1:H1')
    ws['A1'].value = clean(head.get("company_name")) or "SHARMA ENTERPRISES" # Default fallback
    ws['A1'].font = Font(size=22, bold=True)
    ws['A1'].fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
    ws['A1'].alignment = center
    
    # 2. SUBTEXT (A2:H2)
    ws.merge_cells('A2:H2')
    subtext = clean(head.get("company_subtext"))
    if clean(head.get("gstin")): subtext += f" | GSTIN: {clean(head.get('gstin'))}"
    ws['A2'].value = subtext
    ws['A2'].alignment = center
    ws['A2'].border = box_border

    # 3. BUYER & INVOICE DETAILS (Rows 3-4)
    ws.merge_cells('A3:D3'); ws['A3'].value = f"To: {clean(head.get('buyer_name'))}"; ws['A3'].font = Font(bold=True); ws['A3'].border = box_border
    ws.merge_cells('E3:H3'); ws['E3'].value = f"Inv No: {clean(head.get('invoice_no'))}"; ws['E3'].font = Font(bold=True); ws['E3'].alignment = center; ws['E3'].border = box_border
    
    ws.merge_cells('A4:D4'); ws['A4'].value = clean(head.get('buyer_address')) or "Address"; ws['A4'].border = box_border
    ws.merge_cells('E4:H4'); ws['E4'].value = f"Date: {clean(head.get('date'))}"; ws['E4'].alignment = center; ws['E4'].border = box_border

    # 4. BANK DETAILS (Row 5 - FORCED DRAW)
    curr_row = 5
    bank = head.get("bank_details", {})
    # Always draw the bank row structure even if empty
    ws.merge_cells(f'A{curr_row}:B{curr_row}'); ws[f'A{curr_row}'].value = "Bank Details:"; ws[f'A{curr_row}'].font = Font(bold=True); ws[f'A{curr_row}'].border = box_border
    ws.merge_cells(f'C{curr_row}:H{curr_row}')
    ws[f'C{curr_row}'].value = f"{clean(bank.get('bank_name'))} | A/c: {clean(bank.get('acc_no'))} | IFSC: {clean(bank.get('ifsc'))}"
    ws[f'C{curr_row}'].border = box_border
    curr_row += 1

    # 5. TABLE HEADERS (Row 6)
    headers = ["S.N.", "Particulars", "HSN/SAC", "Qty", "Rate", "Gross", "Tax %", "Total"]
    widths = [6, 40, 12, 10, 12, 15, 10, 18]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=curr_row, column=i, value=h)
        c.font = Font(bold=True); c.alignment = center; c.border = box_border
        ws.column_dimensions[get_column_letter(i)].width = w
    
    curr = curr_row + 1
    
    # 6. ITEMS LOOP
    for item in items:
        vals = [item.get("sn"), item.get("particulars"), item.get("hsn_sac"), item.get("quantity"), item.get("rate"), item.get("gross_amount"), item.get("tax_rate"), item.get("amount")]
        for i, val in enumerate(vals, 1):
            c = ws.cell(row=curr, column=i, value=clean(val)); c.border = box_border
            if i in [5, 6, 8]: c.alignment = right
            elif i == 2: c.alignment = left
            else: c.alignment = center
        curr += 1

    # 7. TOTAL ROW
    ws.merge_cells(f'A{curr}:G{curr}')
    ws.cell(row=curr, column=1, value="Total Amount (Inc. GST)").alignment = right
    ws.cell(row=curr, column=1).font = Font(bold=True)
    ws.cell(row=curr, column=1).border = box_border
    ws.cell(row=curr, column=8, value=clean(foot.get("total_amount"))).font = Font(bold=True)
    ws.cell(row=curr, column=8).alignment = right
    ws.cell(row=curr, column=8).border = box_border
    curr += 1

    # 8. AMOUNT IN WORDS & SIGNATURE
    ws.merge_cells(f'A{curr}:H{curr}')
    ws.cell(row=curr, column=1, value=f"Amount in Words: {clean(foot.get('amount_in_words'))}")
    ws.cell(row=curr, column=1).font = Font(italic=True, bold=True)
    ws.cell(row=curr, column=1).border = box_border
    curr += 1
    
    ws.merge_cells(f'F{curr}:H{curr}')
    ws.cell(row=curr, column=6, value="Authorized Signature")
    ws.cell(row=curr, column=6).alignment = right
    ws.cell(row=curr, column=6).font = Font(bold=True)

def write_personal_layout(ws, data):
    items, foot = data.get("items", []), data.get("footer", {})
    ws['A1'] = "EXPENSE SHEET"; ws['A1'].font = Font(size=16, bold=True)
    
    headers = ["Description", "Quantity", "Rate", "Amount"]
    widths = [40, 10, 15, 15]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=4, column=i, value=h); c.font = Font(bold=True)
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

# ==============================================================================
# 7. ROUTES & RENDER BINDING
# ==============================================================================
@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response

def generate_excel(data, save_path):
    wb = Workbook(); ws = wb.active
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
    
    if not file and not prompt_text:
        return jsonify({"error": "No input"}), 400
    
    # 1. Image Logic (Required for Pillow)
    original_name = None
    if file:
        original_name = secure_filename(file.filename)
        img_path = os.path.join(UPLOAD_DIR, original_name)
        file.save(img_path)
    else:
        img_path = os.path.join(UPLOAD_DIR, "temp_blank.png")
        Image.new('RGB', (500, 500), color='white').save(img_path)

    try:
        # 2. Extract Data
        data = parse_invoice_vision(img_path, prompt_text)
        
        # 3. Generate Excel with STRICT LAYOUT
        save_path = os.path.join(UPLOAD_DIR, f"{original_name or 'Expense_Data'}.xlsx")
        generate_excel(data, save_path)
        global LAST_GENERATED_FILE; LAST_GENERATED_FILE = save_path
        
        if not current_user.is_authenticated: session['usage_count'] = session.get('usage_count', 0) + 1
        return jsonify({"status": "ok", "filename": f"{original_name or 'Expense_Data'}.xlsx", "trials_left": 3-session.get('usage_count', 0)})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route("/download")
def download():
    path = os.path.join(UPLOAD_DIR, secure_filename(request.args.get('filename')))
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
