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

# STRICTLY GITHUB MODELS (GPT-4o)
API_KEY = "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"
MODEL = "gpt-4o" 
API_URL = "https://models.inference.ai.azure.com/chat/completions"

# File System Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
LAST_GENERATED_FILE = None

# Ensure upload directory exists
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Initialize Flask App
app = Flask(__name__)

# SECURITY CONFIGURATION
# Using a fixed secret key to ensure sessions persist across restarts
app.config['SECRET_KEY'] = 'FINAL_FULL_CODE_RESTORED_V13' 
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['PERMANENT_SESSION_LIFETIME'] = datetime.timedelta(days=30) # Sessions last 30 days

# ==============================================================================
# 2. DATABASE & AUTHENTICATION MODELS
# ==============================================================================

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login' # Redirect to 'login' if user is not auth

# User Table Model
class User(UserMixin, db.Model):
    """
    User database model.
    Stores name, email, hashed password, and gender.
    """
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150))
    email = db.Column(db.String(150), unique=True)
    password = db.Column(db.String(150))
    gender = db.Column(db.String(50))

# --- NEW: History Model ---
class History(db.Model):
    """
    Stores user activity history.
    Links to the User table via user_id.
    """
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    filename = db.Column(db.String(300))
    prompt = db.Column(db.Text)
    timestamp = db.Column(db.DateTime, default=datetime.datetime.utcnow)

# User Loader for Flask-Login
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Create Database Tables (Run once)
with app.app_context():
    db.create_all()

# ==============================================================================
# 3. UTILITY FUNCTIONS
# ==============================================================================

def clean(v):
    """
    Cleans up cell values to remove None, Null, or Empty strings.
    Returns a clean string.
    This prevents 'None' from appearing in the Excel sheet.
    """
    if v is None: 
        return ""
    s = str(v).strip()
    if s.lower() in ["null", "none", "n/a", "", "[]", "{}"]: 
        return ""
    return s

def encode_image(image_path):
    """
    Encodes an image file to Base64 string for API transmission.
    Required for the Vision API to see the invoice.
    """
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def extract_number(value):
    """
    Extracts the first valid number from a string.
    Crucial for handling units like '50 Mtr', '10 Boxes', or currency symbols.
    
    UPDATED: Handles Negative Numbers for discounts (e.g. -4500).
    """
    if not value: 
        return 0.0
    
    # Regex to capture negative or positive float/int
    # Looks for optional minus sign, then digits, optional decimal
    matches = re.findall(r"(-?\d+(?:\.\d+)?)", str(value).replace(",", ""))
    
    if matches:
        try:
            return float(matches[0])
        except ValueError:
            return 0.0
    return 0.0

def generate_error_excel(error_msg, save_path):
    """
    Generates a valid Excel file containing the error message.
    This prevents the user from receiving a corrupt file if the script crashes.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Error Log"
    
    # Write Header
    ws['A1'] = "PROCESSING FAILED"
    ws['A1'].font = Font(color="FF0000", size=14, bold=True)
    
    # Write Error Message
    ws['A2'] = str(error_msg)
    ws.column_dimensions['A'].width = 60
    
    wb.save(save_path)

# ==============================================================================
# 4. MATH ENGINE (SHARMA, BLUEORBIT, JAN AUSHADI LOGIC)
# ==============================================================================

def recalculate_math(data):
    """
    The core logic to fix amounts, taxes, and totals.
    - Handles Smart Tax Slab Correction (9% -> 18%)
    - Handles Item Level Discounts (Jan Aushadi)
    - Handles Negative Line Items
    - Calculates Gross and Net dynamically
    
    This function ensures that even if the AI misreads a number, 
    the final totals in Excel are mathematically consistent.
    """
    items = data.get("items", [])
    footer = data.get("footer", {})
    layout = data.get("layout", "business")
    
    # SAFETY CHECK: If layout is 'personal', we skip complex tax logic
    is_personal_mode = (layout == "personal")

    running_total = 0.0
    
    # Standard GST slabs for validation (0, 5, 12, 18, 28)
    STANDARD_RATES = [0.0, 5.0, 12.0, 18.0, 28.0]

    # --------------------------------------------------------------------------
    # 4.1 DETECT GLOBAL TAX (Fix for Sharma Ent)
    # --------------------------------------------------------------------------
    global_tax_pct = 0.0
    
    if not is_personal_mode:
        tax_summary = str(footer.get("tax_summary") or "")
        
        # Find numbers like 9, 9.0, 18 in the tax summary
        global_rates = re.findall(r"(\d+(?:\.\d+)?)", tax_summary)
        
        if global_rates:
            # Filter out ridiculously high numbers (amounts) to keep only rates
            raw_nums = [float(r) for r in global_rates if float(r) <= 50]
            
            if raw_nums:
                s = sum(raw_nums) # e.g. 9+9 = 18
                m = max(raw_nums) # e.g. 18 or 9
                
                # Logic: If Sum is a standard rate, use Sum. Else use Max.
                if any(abs(s - x) < 0.1 for x in STANDARD_RATES):
                    global_tax_pct = s
                else:
                    global_tax_pct = m
        
        # *** SMART SLAB CORRECTION (CRITICAL FIX FOR SHARMA) ***
        # If we detected "9%", it's definitely "9% CGST + 9% SGST" -> 18% Total.
        # This fixes the 14715 vs 15930 error logic.
        if abs(global_tax_pct - 9.0) < 0.1: 
            global_tax_pct = 18.0
        elif abs(global_tax_pct - 6.0) < 0.1: 
            global_tax_pct = 12.0
        elif abs(global_tax_pct - 2.5) < 0.1: 
            global_tax_pct = 5.0
        elif abs(global_tax_pct - 14.0) < 0.1: 
            global_tax_pct = 28.0

    # --------------------------------------------------------------------------
    # 4.2 PROCESS ITEMS ROW BY ROW
    # --------------------------------------------------------------------------
    for item in items:
        # A. Clean Inputs (Using Regex for precision)
        qty = extract_number(item.get("quantity"))
        rate = extract_number(item.get("rate"))
        # New: Check for Item Level Discount % (Jan Aushadi case)
        disc_pct = extract_number(item.get("discount_percent"))
        
        desc = str(item.get("particulars") or "").lower()

        # LOGIC FIX FOR "DISCOUNTS" / "ADJUSTMENTS"
        # If the description implies a reduction, force negative rate
        if ("discount" in desc or "adjustment" in desc or "less" in desc) and rate > 0:
            rate = -1 * abs(rate)

        # Default Qty to 1 only if we have a Rate but no Qty (Phantom Item protection)
        if qty == 0 and rate != 0: 
            qty = 1.0

        # B. Base & Gross Amount Logic
        if rate != 0:
            gross_amount = qty * rate
        else:
            # If no rate, look at amount directly
            gross_amount = extract_number(item.get("amount"))
            # Apply discount logic to base amount if rate was missing
            if ("discount" in desc or "adjustment" in desc) and gross_amount > 0:
                gross_amount = -1 * abs(gross_amount)

        # *** JAN AUSHADI FIX: Calculate Discount Amount ***
        discount_amount = 0.0
        if disc_pct > 0:
            # Calculate discount value
            discount_amount = gross_amount * (disc_pct / 100.0)
            
        # Net Taxable Value (Gross - Discount)
        taxable_value = gross_amount - discount_amount

        # C. Tax Inheritance Logic (Complex Logic Restored)
        applicable_tax_pct = 0.0
        display_pct = 0.0
        
        if not is_personal_mode:
            item_tax_str = str(item.get("tax_rate") or "")
            item_tax_nums = re.findall(r"(\d+(?:\.\d+)?)", item_tax_str)
            
            if item_tax_nums:
                # Check item specific tax
                nums = [float(r) for r in item_tax_nums if float(r) <= 100]
                if nums:
                    s = sum(nums)
                    m = max(nums)
                    if any(abs(s - x) < 0.1 for x in STANDARD_RATES):
                        applicable_tax_pct = s
                    else:
                        applicable_tax_pct = m
            
            # *** FORCE INHERITANCE ***
            # Inherit global tax if item has none.
            # Also apply slab correction to item level if needed (e.g. item says 9%)
            if applicable_tax_pct == 0 and global_tax_pct > 0:
                applicable_tax_pct = global_tax_pct
            
            # Apply same 9% -> 18% fix for items
            if abs(applicable_tax_pct - 9.0) < 0.1: 
                applicable_tax_pct = 18.0
            elif abs(applicable_tax_pct - 6.0) < 0.1: 
                applicable_tax_pct = 12.0

        # D. Calculate Final Amount (Base + Tax)
        calc_factor = 0.0
        if applicable_tax_pct > 0:
             if applicable_tax_pct < 1.0: 
                 display_pct = applicable_tax_pct * 100
                 calc_factor = applicable_tax_pct
             else:
                 display_pct = applicable_tax_pct
                 calc_factor = applicable_tax_pct / 100.0

        # Calculate Tax Amount on Taxable Value
        tax_amount_val = taxable_value * calc_factor
        
        # Final Item Total
        final_item_total = taxable_value + tax_amount_val

        # E. Update Item Data with NEW FIELDS
        item["quantity"] = qty
        item["rate"] = rate
        item["gross_amount"] = round(gross_amount, 2)
        item["discount_amount"] = round(discount_amount, 2)
        item["amount"] = round(final_item_total, 2)
        
        if is_personal_mode:
             item["amount"] = round(taxable_value, 2) # No tax in personal
        else:
             # Force the Tax Column to show the percentage used (e.g. "18%")
             item["tax_rate"] = f"{int(display_pct)}%" if display_pct > 0 else "0%"
        
        running_total += item["amount"]

    # F. Force Footer Total
    footer["total_amount"] = round(running_total, 2)
    
    data["items"] = items
    data["footer"] = footer
    return data

# ==============================================================================
# 5. LLM / AI PARSING LOGIC
# ==============================================================================

def parse_invoice_vision(image_path, user_instruction=""):
    base64_image = encode_image(image_path)
    
    # PROMPT: Explicitly asks for Discount Percent and Handles Jan Aushadi
    prompt = f"""
    Extract data into JSON. USER INSTRUCTION: "{user_instruction}"
    
    CRITICAL LAYOUT RULES:
    1. **PERSONAL MODE**: 
       - If image is a list, note, or prompt like "Mr Mehta..." -> Set "layout": "personal".
       - In Personal Mode, DO NOT extract taxes.
    2. **BUSINESS MODE**: 
       - Only use this if "GSTIN" or "Tax Invoice" is present.
    
    CRITICAL EXTRACTION RULES:
    1. **CONTACT INFO**:
       - Look for **Phone Numbers** (10 digits) and **Emails**.
       - Extract to 'phone' and 'email' fields.
    2. **MATH & DISCOUNTS**:
       - **Item Discount**: Look for "Disc%" or "Discount". Extract % to 'discount_percent'.
       - **Service goodwill adjustment**: Treat as Negative Rate.
    3. **TAX SUMMARY**:
       - Extract TAX RATE (e.g. "18%") into 'tax_summary'.
    
    JSON STRUCTURE:
    {{
      "layout": "business" or "personal",
      "header": {{ 
         "company_name": null, "company_subtext": null, "gstin": null, "msme_no": null,
         "buyer_name": null, "buyer_address": null, 
         "date": null, "invoice_no": null, "customer_id": null,
         "challan_no": null, "challan_date": null, "eway_bill_no": null,
         "transport_id": null, "transport_phone": null,
         "bank_details": {{ "bank_name": null, "acc_no": null, "ifsc": null }}
      }},
      "items": [ 
        {{ 
           "sn": "1", 
           "particulars": null, 
           "phone": null, 
           "email": null, 
           "hsn_sac": null, 
           "quantity": null, 
           "rate": null, 
           "per": null, 
           "discount_percent": null,
           "amount": null, 
           "tax_rate": null 
        }} 
      ],
      "footer": {{ 
         "tax_summary": null,
         "total_amount": null, 
         "amount_in_words": null 
      }}
    }}
    """
    
    payload = {
        "model": MODEL,
        "messages": [
            {
                "role": "user", 
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }
        ]
    }
    
    headers = {
        "Authorization": f"Bearer {API_KEY}", 
        "Content-Type": "application/json"
    }

    try:
        r = requests.post(API_URL, headers=headers, json=payload, timeout=60)
        
        if r.status_code != 200:
            raise Exception(f"API Error ({r.status_code}): {r.text}")

        raw = r.json()["choices"][0]["message"]["content"]
        raw = raw.replace("```json", "").replace("```", "")
        
        start_idx = raw.find("{")
        end_idx = raw.rfind("}") + 1
        
        if start_idx == -1: 
            raise Exception("AI returned text but no JSON structure found.")
            
        json_str = raw[start_idx:end_idx]
        
        # --- FIX FOR "INVALID CONTROL CHARACTER" (JSON CRASH FIX) ---
        json_str = re.sub(r'[\x00-\x09\x0b-\x1f\x7f]', '', json_str)
        
        try:
            data = json.loads(json_str, strict=False)
        except json.JSONDecodeError:
            json_str = json_str.replace('\n', ' ').replace('\r', '')
            data = json.loads(json_str, strict=False)

        return recalculate_math(data)

    except Exception as e:
        print(f"LLM Logic Error: {e}")
        raise e 

# ==============================================================================
# 6. EXCEL LAYOUTS (DETAILED & CONDITIONAL)
# ==============================================================================

def write_business_layout(ws, data):
    """
    Writes the Business Invoice Layout.
    Includes full borders, merging, and specific alignment rules.
    """
    head = data.get("header", {})
    items = data.get("items", [])
    foot = data.get("footer", {})
    
    # --------------------------------------------------------------------------
    # 6.1 SETUP COLUMNS & HEADERS
    # --------------------------------------------------------------------------
    has_hsn = any(clean(item.get("hsn_sac")) for item in items)
    # Check if we have any valid discount amounts calculated
    has_disc = any(item.get("discount_amount", 0) > 0 for item in items)
    
    # DYNAMIC HEADER CONSTRUCTION
    headers = ["S.N.", "Particulars"]
    widths = [6, 40]
    keys = ["sn", "particulars"]
    
    if has_hsn:
        headers.append("HSN/SAC"); widths.append(12); keys.append("hsn_sac")
        
    headers.extend(["Qty", "Rate"])
    widths.extend([10, 12])
    keys.extend(["quantity", "rate"])
    
    # NEW: Add Gross Amount & Discount ONLY if Discount Exists
    if has_disc:
        headers.append("Gross Amt")
        widths.append(15)
        keys.append("gross_amount")
        
        headers.append("Discount")
        widths.append(12)
        keys.append("discount_amount")
        
    headers.extend(["Tax %", "Amount (Inc. Tax)"])
    widths.extend([10, 18])
    keys.extend(["tax_rate", "amount"])

    num_cols = len(headers)
    last_col = get_column_letter(num_cols)
    
    # Define Styles
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    
    thick_side = Side(style='medium', color='000000')
    thin_side = Side(style='thin', color='000000')
    box_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)
    
    def style_range(r, border=None, fill=None, font=None, align=None):
        rows = ws[r]
        if not isinstance(rows, tuple): rows = (rows,)
        for row in rows:
            for c in row:
                if border: c.border = border
                if fill: c.fill = fill
                if font: c.font = font
                if align: c.alignment = align

    # --------------------------------------------------------------------------
    # 6.2 WRITE TITLE & COMPANY INFO
    # --------------------------------------------------------------------------
    ws.merge_cells(f'A1:{last_col}1')
    ws['A1'].value = clean(head.get("company_name")) or "INVOICE"
    
    orange_fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
    title_font = Font(name='Calibri', size=22, bold=True)
    
    style_range(f'A1:{last_col}1', border=box_border, fill=orange_fill, font=title_font, align=center)

    # Subtext (Address, GSTIN)
    subtext = clean(head.get("company_subtext"))
    if clean(head.get("gstin")): 
        subtext += f" | GSTIN: {clean(head.get('gstin'))}"
    if clean(head.get("msme_no")): 
        subtext += f" | MSME: {clean(head.get('msme_no'))}"
    
    ws.merge_cells(f'A2:{last_col}2')
    ws['A2'].value = subtext
    style_range(f'A2:{last_col}2', border=box_border, font=Font(size=10), align=center)
    
    if len(subtext) > 60: 
        ws.row_dimensions[2].height = 40

    # --------------------------------------------------------------------------
    # 6.3 WRITE BUYER & INVOICE DETAILS
    # --------------------------------------------------------------------------
    mid_col_idx = num_cols // 2 + 1
    mid_col = get_column_letter(mid_col_idx)
    prev_col = get_column_letter(mid_col_idx - 1)

    # Row 3: Buyer Name | Invoice No
    ws.merge_cells(f'A3:{prev_col}3')
    ws['A3'].value = f"To: {clean(head.get('buyer_name'))}"
    style_range(f'A3:{prev_col}3', border=box_border, align=left, font=Font(bold=True))
    
    ws.merge_cells(f'{mid_col}3:{last_col}3')
    inv_no = clean(head.get('invoice_no'))
    ws[f'{mid_col}3'].value = f"Inv No: {inv_no}" if inv_no else ""
    style_range(f'{mid_col}3:{last_col}3', border=box_border, align=center, font=Font(bold=True))

    # Row 4: Buyer Address | Date & Customer ID
    ws.merge_cells(f'A4:{prev_col}4')
    ws['A4'].value = clean(head.get('buyer_address'))
    style_range(f'A4:{prev_col}4', border=box_border, align=left)

    ws.merge_cells(f'{mid_col}4:{last_col}4')
    date_txt = clean(head.get('date'))
    cust_id = clean(head.get('customer_id'))
    right_text = f"Date: {date_txt}"
    if cust_id: 
        right_text += f"\nCust ID: {cust_id}"
        
    ws[f'{mid_col}4'].value = right_text
    style_range(f'{mid_col}4:{last_col}4', border=box_border, align=center)
    ws.row_dimensions[4].height = 30

    # --------------------------------------------------------------------------
    # 6.4 WRITE OPTIONAL FIELDS (Transport, Bank)
    # --------------------------------------------------------------------------
    curr_row = 5
    optional_fields = [
        ("Challan No", clean(head.get("challan_no"))),
        ("Challan Date", clean(head.get("challan_date"))),
        ("E-Way Bill No", clean(head.get("eway_bill_no"))),
        ("Transport ID", clean(head.get("transport_id"))),
        ("Transport Phone", clean(head.get("transport_phone"))),
    ]
    
    # Add Bank Details
    bank = head.get("bank_details", {})
    if clean(bank.get("acc_no")):
         bank_str = f"{clean(bank.get('bank_name'))} | A/c: {clean(bank.get('acc_no'))} | IFSC: {clean(bank.get('ifsc'))}"
         optional_fields.append(("Bank Details", bank_str))

    for label, value in optional_fields:
        if value:
            ws.merge_cells(f'A{curr_row}:B{curr_row}')
            ws.cell(row=curr_row, column=1, value=label + ":")
            style_range(f'A{curr_row}:B{curr_row}', border=box_border, align=left, font=Font(bold=True))
            
            ws.merge_cells(f'C{curr_row}:{last_col}{curr_row}')
            ws.cell(row=curr_row, column=3, value=value)
            style_range(f'C{curr_row}:{last_col}{curr_row}', border=box_border, align=left)
            curr_row += 1

    # --------------------------------------------------------------------------
    # 6.5 WRITE TABLE HEADERS
    # --------------------------------------------------------------------------
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=curr_row, column=i, value=h)
        c.border = box_border
        c.font = Font(bold=True)
        c.alignment = center
        ws.column_dimensions[get_column_letter(i)].width = w
    
    table_start = curr_row
    curr = curr_row + 1
    
    # --------------------------------------------------------------------------
    # 6.6 WRITE LINE ITEMS
    # --------------------------------------------------------------------------
    for item in items:
        for i, key in enumerate(keys, 1):
            val = clean(item.get(key))
            c = ws.cell(row=curr, column=i, value=val)
            c.border = box_border
            
            # Alignments
            if key == "particulars": 
                c.alignment = left
            elif key in ["amount", "gross_amount", "discount_amount"]: 
                c.alignment = right
            else: 
                c.alignment = center
        curr += 1
    
    # Fill empty rows
    fill_to = table_start + 10
    while curr <= fill_to:
        for i in range(1, num_cols + 1): 
            ws.cell(row=curr, column=i).border = box_border
        curr += 1

    # --------------------------------------------------------------------------
    # 6.7 WRITE TOTALS & FOOTER
    # --------------------------------------------------------------------------
    # Total Label
    total_label_end = get_column_letter(num_cols - 1)
    ws.merge_cells(f'A{curr}:{total_label_end}{curr}')
    ws.cell(row=curr, column=1, value="Total Amount (Inc. GST)").alignment = right
    style_range(f'A{curr}:{total_label_end}{curr}', border=box_border, font=Font(bold=True), align=right)
    
    # Total Value
    c = ws.cell(row=curr, column=num_cols, value=clean(foot.get("total_amount")))
    c.border = box_border
    c.font = Font(bold=True)
    c.alignment = right
    curr += 1

    # Amount in Words
    ws.merge_cells(f'A{curr}:{last_col}{curr}')
    word_text = clean(foot.get("amount_in_words"))
    if word_text: 
        word_text = f"Amount in Words: {word_text}"
    
    ws.cell(row=curr, column=1, value=word_text)
    style_range(f'A{curr}:{last_col}{curr}', border=box_border, align=left, font=Font(bold=True, italic=True))
    curr += 1

    # Signature Block
    ws.merge_cells(f'A{curr}:C{curr}')
    style_range(f'A{curr}:C{curr}', border=box_border)
    
    sig_start = get_column_letter(4)
    ws.merge_cells(f'{sig_start}{curr}:{last_col}{curr}')
    ws.cell(row=curr, column=4, value="Authorized Signature")
    style_range(f'{sig_start}{curr}:{last_col}{curr}', border=box_border, align=Alignment(horizontal='right', vertical='bottom'), font=Font(bold=True))
    ws.row_dimensions[curr].height = 40

def write_personal_layout(ws, data):
    """
    Writes the Personal / Data Entry Layout.
    Dynamically adds Phone/Email AND GROSS/DISCOUNT columns ONLY if needed.
    """
    head = data.get("header", {})
    items = data.get("items", [])
    foot = data.get("footer", {})
    
    # DYNAMIC COLUMN DETECTION
    has_qty = any(clean(item.get("quantity")) for item in items)
    has_phone = any(clean(item.get("phone")) for item in items)
    has_email = any(clean(item.get("email")) for item in items)
    
    # CRITICAL: Only check for discounts > 0 to enable columns
    has_disc = any(item.get("discount_amount", 0) > 0 for item in items)

    # Build Header List Dynamically
    headers = ["Description"]
    cols = ["particulars"]
    widths = [40]

    if has_phone:
        headers.append("Phone")
        cols.append("phone")
        widths.append(15)
    
    if has_email:
        headers.append("Email")
        cols.append("email")
        widths.append(25)

    if has_qty:
        # If Discount Exists -> Detailed Breakdown (Gross, Disc, Net)
        if has_disc:
            headers.extend(["Quantity", "Rate", "Gross Amt", "Discount", "Net Amount"])
            cols.extend(["quantity", "rate", "gross_amount", "discount_amount", "amount"])
            widths.extend([10, 10, 15, 12, 18])
        else:
            # No Discount -> Simple Layout
            headers.extend(["Quantity", "Rate", "Amount"])
            cols.extend(["quantity", "rate", "amount"])
            widths.extend([10, 10, 18])
    else:
        headers.append("Amount")
        cols.append("amount")
        widths.append(20)

    # Style Header
    ws['A1'] = "EXPENSE SHEET"
    ws['A1'].font = Font(size=16, bold=True, color="444444")
    
    if head.get('date'): 
        ws['A2'] = f"Date: {clean(head.get('date'))}"
        ws['A2'].font = Font(italic=True)

    # Write Headers
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(bold=True)
        c.border = Border(bottom=Side(style='thin'))
        ws.column_dimensions[get_column_letter(i)].width = w

    # Write Data
    curr = 5
    for item in items:
        for i, key in enumerate(cols, 1):
            val = clean(item.get(key))
            c = ws.cell(row=curr, column=i, value=val)
            
            # Alignments
            if key in ["amount", "rate", "quantity", "gross_amount", "discount_amount"]: 
                c.alignment = Alignment(horizontal='right')
            elif key == "phone":
                c.alignment = Alignment(horizontal='center')
        curr += 1

    # Total Row
    curr += 1
    total_col_idx = len(cols)
    ws.cell(row=curr, column=total_col_idx - 1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=curr, column=total_col_idx - 1).alignment = Alignment(horizontal='right')
    
    ws.cell(row=curr, column=total_col_idx, value=foot.get("total_amount")).font = Font(bold=True)
    ws.cell(row=curr, column=total_col_idx).alignment = Alignment(horizontal='right')

# ==============================================================================
# 7. ROUTES (HISTORY, PROFILE, RESTORE & DOWNLOAD)
# ==============================================================================

# *** CACHE BUSTER TO FIX BACK BUTTON ISSUE ***
# This forces the browser to check with the server every time a page is loaded.
# This ensures that when you click "Back" to Home, you get the logged-in version (Avatar),
# not the cached guest version (Login buttons).
@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

def generate_excel(data, save_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.sheet_view.showGridLines = True 

    if not data: 
        wb.save(save_path)
        return

    layout = data.get("layout", "business")
    try:
        if layout == "personal": 
            write_personal_layout(ws, data)
        else: 
            write_business_layout(ws, data)
    except Exception as e:
        ws['A1'] = f"Layout Error: {e}"
    
    wb.save(save_path)

@app.route("/")
def home(): 
    # PASS USER TO TEMPLATE SO INDEX.HTML CAN SHOW AVATAR
    return render_template("index.html", user=current_user)

# --- AUTH ROUTES (WITH FIXES) ---

@app.route("/login", methods=['GET', 'POST'])
def login():
    # *** NAVIGATION FIX ***
    # If user is already logged in and hits Back to reach Login page,
    # Redirect BACK TO HOME immediately. 
    # This creates the effect: Input -> Back -> Home (Skipping Login)
    if current_user.is_authenticated:
        return redirect(url_for('home'))

    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        
        user = User.query.filter_by(email=email).first()
        
        if user and check_password_hash(user.password, password):
            login_user(user)
            # Do NOT reset 'usage_count' to allow persistent trial tracking
            return redirect(url_for('input_page'))
        flash('Login failed. Check email and password.', 'danger')
    return render_template('login.html')

@app.route("/signup", methods=['GET', 'POST'])
def signup():
    # *** NAVIGATION FIX ***
    if current_user.is_authenticated:
        return redirect(url_for('home'))

    if request.method == 'POST':
        name = request.form.get('name')
        email = request.form.get('email')
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')
        gender = request.form.get('gender')

        if password != confirm_password:
            flash('Passwords do not match!', 'danger')
            return redirect(url_for('signup'))
        
        user_exists = User.query.filter_by(email=email).first()
        if user_exists:
            flash('Email already exists.', 'warning')
            return redirect(url_for('signup'))

        # Create new user
        new_user = User(
            name=name, 
            email=email, 
            password=generate_password_hash(password, method='scrypt'), 
            gender=gender
        )
        db.session.add(new_user)
        db.session.commit()
        
        # SUCCESS MESSAGE CONFIRMATION
        flash('Account created successfully! You can now log in.', 'success')
        return redirect(url_for('login'))
        
    return render_template('signup.html')

@app.route("/logout")
@login_required
def logout():
    # Capture usage before logout clears session
    old_usage = session.get('usage_count', 0)
    logout_user()
    # Restore usage to anonymous session
    session['usage_count'] = old_usage
    return redirect(url_for('home'))

# --- APP ROUTES (WITH 3-TRIAL LIMIT) ---

@app.route("/input")
def input_page(): 
    # Remove @login_required. Check status manually.
    if not current_user.is_authenticated:
        session.permanent = True
        
        # Ensures usage_count is explicitly 0 if missing (Fixes "0 Trials" bug)
        if 'usage_count' not in session:
            session['usage_count'] = 0
            
        usage = session.get('usage_count', 0)
        return render_template("input.html", user=None, trials_left=3-usage)
    return render_template("input.html", user=current_user, trials_left=None)

# --- NEW: Routes for Profile and History Data ---
@app.route('/get_profile')
@login_required
def get_profile():
    return jsonify({
        "name": current_user.name,
        "email": current_user.email,
        "gender": current_user.gender
    })

@app.route('/get_history')
@login_required
def get_history():
    # Fetch history reverse ordered by time
    items = History.query.filter_by(user_id=current_user.id).order_by(History.timestamp.desc()).all()
    history_list = []
    for item in items:
        history_list.append({
            "id": item.id,
            "filename": item.filename,
            "prompt": item.prompt,
            "date": item.timestamp.strftime("%d %b %Y, %I:%M %p")
        })
    return jsonify(history_list)

@app.route("/process", methods=["POST"])
def process():
    current_trials = None 

    # 1. CHECK TRIALS BEFORE PROCESSING
    if not current_user.is_authenticated:
        # Ensures usage_count is explicitly 0 if missing (Fixes "0 Trials" bug)
        if 'usage_count' not in session:
            session['usage_count'] = 0
            
        usage = session.get('usage_count', 0)
        if usage >= 3:
             # This triggers the specific 403 error message in JS
             return jsonify({"error": "3 trials ended now do log in"}), 403

    global LAST_GENERATED_FILE
    
    file = request.files.get("image")
    prompt_text = request.form.get("prompt", "")
    restored_filename = request.form.get("restored_filename", "") # Check if restoring old file

    # FILE HANDLING: New Upload OR Restore Old File
    if file:
        original_name = secure_filename(file.filename) or "image.png"
        img_path = os.path.join(UPLOAD_DIR, original_name)
        file.save(img_path)
    elif restored_filename:
        # Use the file already on server
        original_name = restored_filename
        img_path = os.path.join(UPLOAD_DIR, restored_filename)
        if not os.path.exists(img_path): 
            return jsonify({"error": "Original file not found on server. Please re-upload."}), 404
    else:
        if not prompt_text: 
            return jsonify({"error": "No input"}), 400
        original_name = None
        img_path = os.path.join(UPLOAD_DIR, "temp_blank.png")
        Image.new('RGB', (500, 500), color='white').save(img_path)

    excel_name = "Expense_Data.xlsx"
    if original_name: 
        excel_name = f"{os.path.splitext(original_name)[0]}.xlsx"
        
    save_path = os.path.join(UPLOAD_DIR, excel_name)
    
    try:
        parsed_data = parse_invoice_vision(img_path, user_instruction=prompt_text)
        generate_excel(parsed_data, save_path)
        
        # SAVE TO HISTORY (If Logged In)
        if current_user.is_authenticated:
            new_hist = History(user_id=current_user.id, filename=original_name, prompt=prompt_text)
            db.session.add(new_hist)
            db.session.commit()
        else:
            # 2. INCREMENT TRIAL COUNT ONLY IF SUCCESSFUL (THE FIX)
            session['usage_count'] = session.get('usage_count', 0) + 1
            current_trials = 3 - session['usage_count']
            
    except Exception as e:
        print(f"PROCESS ERROR: {e}")
        generate_error_excel(str(e), save_path)
        # Do not increment trial count here
    
    LAST_GENERATED_FILE = save_path
    
    # Return trials_left so frontend updates instantly
    return jsonify({
        "status": "ok", 
        "filename": excel_name, 
        "trials_left": current_trials
    })

@app.route("/download")
def download():
    # Allow download even if trial just expired, as they just made it.
    requested_name = request.args.get('filename')
    try:
        if requested_name:
            path = os.path.join(UPLOAD_DIR, secure_filename(requested_name))
            if os.path.exists(path): 
                return send_file(path, as_attachment=True, download_name=requested_name)
        
        if LAST_GENERATED_FILE and os.path.exists(LAST_GENERATED_FILE):
            return send_file(LAST_GENERATED_FILE, as_attachment=True, download_name=os.path.basename(LAST_GENERATED_FILE))
            
        return "No file found", 404
    except Exception as e:
        return f"Download Error: {e}", 500

if __name__ == "__main__":
    app.run(debug=True)