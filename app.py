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
app.config['SECRET_KEY'] = 'FINAL_VISUAL_FIX_V19_ROW_HEIGHTS' 
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
# 4. AI PARSING LOGIC (ENHANCED ADDRESS EXTRACTION)
# ==============================================================================
def parse_invoice_vision(image_path, user_instruction=""):
    base64_img = encode_image(image_path)
    
    # PROMPT FIX: Explicitly asking for the address line "Immediately below title"
    prompt = f"""
    Extract data into JSON. USER INSTRUCTION: "{user_instruction}"
    
    RULES:
    1. **PERSONAL**: If input is text only, set "layout": "personal".
    2. **BUSINESS**: If image, set "layout": "business".
       - **COMPANY NAME**: Topmost large text (e.g. "Sharma Enterprises").
       - **SUBTEXT (Address)**: Capture the full address & GSTIN lines found IMMEDIATELY BELOW the Company Name.
       - **INVOICE NO**: Look for "Inv No" or "Invoice #".
    
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
# 5. EXCEL LAYOUTS (FORCED DIMENSIONS)
# ==============================================================================
def draw_box(ws, cell_range, value, font=None, align=None, fill=None, border=None):
    """ Helper to force draw merged cells with styles """
    ws.merge_cells(cell_range)
    top_left = ws[cell_range.split(':')[0]]
    top_left.value = value
    if font: top_left.font = font
    if align: top_left.alignment = align
    if fill: top_left.fill = fill
    if border:
        for row in ws[cell_range]:
            for c in row: c.border = border

def write_business_layout(ws, data):
    head, items, foot = data.get("header", {}), data.get("items", []), data.get("footer", {})
    
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    box_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))
    
    # 1. TITLE (FORCED HEIGHT)
    # Force row height so text isn't squashed
    ws.row_dimensions[1].height = 35 
    comp_name = clean(head.get("company_name")) or "SHARMA ENTERPRISES"
    draw_box(ws, 'A1:H1', comp_name, 
             font=Font(size=22, bold=True), align=center, 
             fill=PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid"), border=box_border)
    
    # 2. SUBTEXT (FORCED HEIGHT)
    ws.row_dimensions[2].height = 20
    subtext = clean(head.get("company_subtext"))
    if clean(head.get("gstin")) and "GSTIN" not in subtext: 
        subtext += f" | GSTIN: {clean(head.get('gstin'))}"
    draw_box(ws, 'A2:H2', subtext, align=center, border=box_border)

    # 3. DETAILS GRID
    ws.row_dimensions[3].height = 20; ws.row_dimensions[4].height = 20
    draw_box(ws, 'A3:D3', f"To: {clean(head.get('buyer_name'))}", font=Font(bold=True), align=left, border=box_border)
    draw_box(ws, 'E3:H3', f"Inv No: {clean(head.get('invoice_no'))}", font=Font(bold=True), align=center, border=box_border)
    draw_box(ws, 'A4:D4', clean(head.get('buyer_address')) or "Address", align=left, border=box_border)
    draw_box(ws, 'E4:H4', f"Date: {clean(head.get('date'))}", align=center, border=box_border)

    # 4. BANK DETAILS
    ws.row_dimensions[5].height = 20
    bank = head.get("bank_details", {})
    draw_box(ws, 'A5:B5', "Bank Details:", font=Font(bold=True), align=left, border=box_border)
    bank_str = f"{clean(bank.get('bank_name'))} | A/c: {clean(bank.get('acc_no'))} | IFSC: {clean(bank.get('ifsc'))}"
    draw_box(ws, 'C5:H5', bank_str, align=left, border=box_border)

    # 5. TABLE HEADERS
    headers = ["S.N.", "Particulars", "HSN/SAC", "Qty", "Rate", "Gross", "Tax %", "Total"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=6, column=i, value=h)
        c.font = Font(bold=True); c.alignment = center; c.border = box_border
        ws.column_dimensions[get_column_letter(i)].width = 15 if i != 2 else 40
    
    # 6. ITEMS
    curr = 7
    for item in items:
        vals = [item.get("sn"), item.get("particulars"), item.get("hsn_sac"), item.get("quantity"), item.get("rate"), item.get("gross_amount"), item.get("tax_rate"), item.get("amount")]
        for i, val in enumerate(vals, 1):
            c = ws.cell(row=curr, column=i, value=clean(val)); c.border = box_border
            c.alignment = right if i in [5,6,8] else (left if i==2 else center)
        curr += 1

    # 7. TOTAL & WORDS
    draw_box(ws, f'A{curr}:G{curr}', "Total Amount (Inc. GST)", font=Font(bold=True), align=right, border=box_border)
    ws.cell(row=curr, column=8, value=clean(foot.get("total_amount"))).font = Font(bold=True)
    ws.cell(row=curr, column=8).border = box_border; ws.cell(row=curr, column=8).alignment = right
    
    curr += 1
    ws.row_dimensions[curr].height = 25
    draw_box(ws, f'A{curr}:H{curr}', f"Amount in Words: {clean(foot.get('amount_in_words'))}", font=Font(italic=True, bold=True), align=left, border=box_border)
    
    # 8. SIGNATURE BOX (FORCED HEIGHT & BORDER)
    curr += 1
    ws.row_dimensions[curr].height = 50 # TALLER FOR SIGNATURE
    
    # Merge cells for signature block
    sig_range = f'F{curr}:H{curr}'
    ws.merge_cells(sig_range)
    sig_cell = ws[f'F{curr}']
    sig_cell.value = "Authorized Signature"
    sig_cell.font = Font(bold=True)
    # Align bottom-right to look like a signature block
    sig_cell.alignment = Alignment(horizontal='right', vertical='bottom')
    
    # Draw border around the signature block
    for row in ws[sig_range]:
        for c in row: c.border = box_border

def write_personal_layout(ws, data):
    items, foot = data.get("items", []), data.get("footer", {})
    ws['A1'] = "EXPENSE SHEET"; ws['A1'].font = Font(size=16, bold=True)
    headers = ["Description", "Quantity", "Rate", "Amount"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h).font = Font(bold=True)
        ws.column_dimensions[get_column_letter(i)].width = 20 if i!=1 else 40
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
# 6. ROUTES
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
    
    img_path = os.path.join(UPLOAD_DIR, secure_filename(file.filename) if file else "prompt.png")
    if file: file.save(img_path)
    else: Image.new('RGB', (500, 500), color='white').save(img_path)

    try:
        data = parse_invoice_vision(img_path, prompt_text)
        wb = Workbook(); ws = wb.active
        if data.get("layout") == "personal": write_personal_layout(ws, data)
        else: write_business_layout(ws, data)
        
        save_path = os.path.join(UPLOAD_DIR, "Data.xlsx"); wb.save(save_path)
        global LAST_GENERATED_FILE; LAST_GENERATED_FILE = save_path
        
        if not current_user.is_authenticated: session['usage_count'] = session.get('usage_count', 0) + 1
        return jsonify({"status": "ok", "filename": "Data.xlsx", "trials_left": 3-session.get('usage_count', 0)})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route("/download")
def download():
    path = os.path.join(UPLOAD_DIR, secure_filename(request.args.get('filename', 'Data.xlsx')))
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
