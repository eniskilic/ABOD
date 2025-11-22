import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch, landscape, A4
from reportlab.lib import colors
import requests
from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_bytes
import pytesseract
from difflib import get_close_matches

# --------------------------------------
# Page Configuration
# --------------------------------------
st.set_page_config(
    page_title="Blanket Order Manager",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --------------------------------------
# Dark Mode Custom CSS Styling (Restored from V10.1)
# --------------------------------------
st.markdown("""
<style>
    /* Dark Mode Base */
    .main { background: #0f1419; color: #e4e6eb; }
    .stApp { background: #0f1419; }
    
    /* Sidebar Dark Styling */
    [data-testid="stSidebar"] { background: #1a1f2e; border-right: 1px solid #2d3748; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] { color: #e4e6eb; }
    
    /* Metric Cards Dark */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #1e2432 0%, #252d3d 100%);
        border: 1px solid #2d3748;
        padding: 25px 20px;
        border-radius: 16px;
        border-left: 3px solid #667eea;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.2);
        border-color: #667eea;
    }
    [data-testid="stMetric"] label { color: #a0aec0 !important; }
    [data-testid="stMetric"] [data-testid="stMetricValue"] { color: #e4e6eb !important; }
    
    /* Buttons Dark */
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 24px;
        font-weight: 600;
        width: 100%;
    }
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.4);
    }
    
    /* File Uploader Dark */
    [data-testid="stFileUploader"] {
        background: #1a1f2e;
        padding: 40px;
        border-radius: 12px;
        border: 2px dashed #2d3748;
    }
    [data-testid="stFileUploader"] label { color: #e4e6eb !important; }
    
    /* Dataframe Dark */
    [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; }
    
    /* Text color overrides */
    p, span, div { color: #cbd5e0; }
    strong { color: #e4e6eb; }
    hr { border-top: 1px solid #2d3748; margin: 40px 0; }
    
    /* Status Indicator */
    .status-indicator {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: #2d3748;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 0.85em;
    }
    .status-dot {
        width: 8px;
        height: 8px;
        background: #48bb78;
        border-radius: 50%;
        animation: pulse 2s infinite;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
</style>
""", unsafe_allow_html=True)

# --------------------------------------
# Configuration & Helpers
# --------------------------------------
AIRTABLE_PAT = "patD9n2LOJRthfGan.b420b57e48143665f27870484e882266bcfd184fa7c96067fbb1ef8c41424fae"
BASE_ID = "appxoNC3r5NSsTP3U"
ORDERS_TABLE = "Orders"
LINE_ITEMS_TABLE = "Order Line Items"

COLOR_TRANSLATIONS = {
    "white": "Blanco", "black": "Negro", "brown": "Marrón", "blue": "Azul",
    "navy": "Azul Marino", "red": "Rojo", "pink": "Rosa", "light pink": "Rosa Claro",
    "hot pink": "Rosa Fucsia", "salmon pink": "Rosa Salmón", "purple": "Morado",
    "lilac": "Lila", "gray": "Gris", "grey": "Gris", "gold": "Dorado",
    "silver": "Plateado", "beige": "Beige", "green": "Verde", "olive": "Verde Oliva",
    "yellow": "Amarillo", "champagne": "Champán"
}

def clean_text(s: str) -> str:
    if not s: return ""
    s = re.sub(r"\(#?[A-Fa-f0-9]{3,6}\)", "", s)
    s = re.sub(r"■|Seller Name|Your Orders|Returning your item:", "", s)
    s = re.sub(r"\(Most popular\)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def translate_thread_color(color):
    if not color: return color
    base = color.strip()
    for eng, esp in COLOR_TRANSLATIONS.items():
        if eng.lower() in base.lower():
            return f"{base} ({esp})"
    return base

def get_bobbin_color(thread_color):
    thread_lower = str(thread_color).lower()
    if 'navy' in thread_lower or 'black' in thread_lower or 'negro' in thread_lower:
        return 'Black Bobbin'
    else:
        return 'White Bobbin'

def draw_checkbox(canvas_obj, x, y, size, is_checked):
    canvas_obj.saveState()
    if is_checked:
        canvas_obj.setStrokeColor(colors.black)
        canvas_obj.setFillColor(colors.black)
        canvas_obj.setLineWidth(2)
        canvas_obj.rect(x, y, size, size, stroke=1, fill=1)
    else:
        canvas_obj.setStrokeColor(colors.black)
        canvas_obj.setLineWidth(2)
        canvas_obj.rect(x, y, size, size, stroke=1, fill=0)
    canvas_obj.restoreState()

# --------------------------------------
# Airtable Functions
# --------------------------------------
def get_existing_order_ids():
    headers = {"Authorization": f"Bearer {AIRTABLE_PAT}", "Content-Type": "application/json"}
    existing = set()
    try:
        url = f"https://api.airtable.com/v0/{BASE_ID}/{ORDERS_TABLE}"
        r = requests.get(url, headers=headers, params={"fields[]": "Order ID"})
        if r.status_code == 200:
            for rec in r.json().get("records", []):
                existing.add(rec["fields"].get("Order ID"))
    except: pass
    return existing

def upload_to_airtable(dataframe):
    headers = {"Authorization": f"Bearer {AIRTABLE_PAT}", "Content-Type": "application/json"}
    existing = get_existing_order_ids()
    unique = dataframe[['Order ID', 'Order Date', 'Buyer Name']].drop_duplicates(subset=['Order ID'])
    new = unique[~unique['Order ID'].isin(existing)]
    
    if len(new) == 0: return 0, 0, []
    
    orders_created, line_items_created, errors = 0, 0, []
    progress = st.progress(0)
    
    for i, (_, row) in enumerate(new.iterrows()):
        try:
            order_payload = {"records": [{"fields": {"Order ID": row['Order ID'], "Order Date": row['Order Date'], "Buyer Name": row['Buyer Name'], "Status": "New"}}]}
            r = requests.post(f"https://api.airtable.com/v0/{BASE_ID}/{ORDERS_TABLE}", headers=headers, json=order_payload)
            if r.status_code == 200:
                oid = r.json()["records"][0]["id"]
                orders_created += 1
                items = dataframe[dataframe['Order ID'] == row['Order ID']]
                for _, item in items.iterrows():
                    li_payload = {"records": [{"fields": {
                        "Order ID": [oid], "Buyer Name": item['Buyer Name'], "Customization Name": item['Customization Name'],
                        "Quantity": int(item['Quantity']), "Blanket Color": item['Blanket Color'], "Thread Color": item['Thread Color'],
                        "Include Beanie": item['Include Beanie'], "Gift Box": item['Gift Box'], "Gift Note": item['Gift Note'],
                        "Gift Message": item['Gift Message'], "Bobbin Color": get_bobbin_color(item['Thread Color']), "Status": "Pending"
                    }}]}
                    r2 = requests.post(f"https://api.airtable.com/v0/{BASE_ID}/{LINE_ITEMS_TABLE}", headers=headers, json=li_payload)
                    if r2.status_code == 200: line_items_created += 1
            else: errors.append(f"Failed Order {row['Order ID']}")
        except Exception as e: errors.append(str(e))
        progress.progress((i+1)/len(new))
    return orders_created, line_items_created, errors
    # --------------------------------------
# PDF Generation Functions (FULL VERBOSE VERSION)
# --------------------------------------
def generate_manufacturing_labels(dataframe):
    buf = BytesIO()
    page_size = landscape((4 * inch, 6 * inch))
    c = canvas.Canvas(buf, pagesize=page_size)
    W, H = page_size
    left = 0.3 * inch
    right = W - 0.3 * inch
    top = H - 0.3 * inch

    for _, row in dataframe.iterrows():
        y = top
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, y, f"Order ID: {row['Order ID']}")
        c.drawRightString(right, y, f"Qty: {row['Quantity']}")
        y -= 0.25 * inch
        
        c.setFont("Helvetica", 14)
        c.drawString(left, y, f"Buyer: {row['Buyer Name']}")
        c.drawRightString(right, y, f"Date: {row['Order Date']}")
        y -= 0.3 * inch

        box_height = 0.7 * inch
        box_y = y - box_height
        c.setStrokeColor(colors.black)
        c.setLineWidth(2)
        c.rect(left, box_y, right - left, box_height, stroke=1, fill=0)
        
        c.setFont("Helvetica-Bold", 16)
        text_y = box_y + box_height - 0.24 * inch
        c.drawString(left + 0.1 * inch, text_y, f"BLANKET COLOR: {str(row['Blanket Color']).upper()}")
        
        text_y -= 0.32 * inch
        c.setFont("Helvetica-BoldOblique", 16)
        c.drawString(left + 0.1 * inch, text_y, f"THREAD COLOR: {row['Thread Color']}")
        
        y = box_y - 0.3 * inch

        c.setFont("Helvetica-Bold", 18)
        c.drawString(left, y, f"★ Name: {row['Customization Name']}")
        y -= 0.4 * inch

        frame_width = (right - left - 0.4 * inch) / 3
        frame_height = 1.1 * inch
        frame_y = y - frame_height
        
        c.setLineWidth(2)
        
        beanie_x = left
        c.rect(beanie_x, frame_y, frame_width, frame_height, stroke=1, fill=0)
        
        checkbox_size = 0.25 * inch
        checkbox_x = beanie_x + (frame_width - checkbox_size) / 2
        checkbox_y = frame_y + frame_height - 0.35 * inch
        is_beanie_checked = (row['Include Beanie'] == "YES")
        draw_checkbox(c, checkbox_x, checkbox_y, checkbox_size, is_beanie_checked)
        
        text_x = beanie_x + frame_width / 2
        text_y = frame_y + frame_height - 0.60 * inch
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(text_x, text_y, "BEANIE")
        
        text_y -= 0.25 * inch
        if row['Include Beanie'] == "YES":
            c.setFont("Helvetica-BoldOblique", 14)
        else:
            c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(text_x, text_y, str(row['Include Beanie']))
        
        gift_box_x = beanie_x + frame_width + 0.2 * inch
        c.rect(gift_box_x, frame_y, frame_width, frame_height, stroke=1, fill=0)
        
        checkbox_x = gift_box_x + (frame_width - checkbox_size) / 2
        is_gift_box_checked = (row['Gift Box'] == "YES")
        draw_checkbox(c, checkbox_x, checkbox_y, checkbox_size, is_gift_box_checked)
        
        text_x = gift_box_x + frame_width / 2
        text_y = frame_y + frame_height - 0.60 * inch
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(text_x, text_y, "GIFT BOX")
        
        text_y -= 0.25 * inch
        if row['Gift Box'] == "YES":
            c.setFont("Helvetica-BoldOblique", 14)
        else:
            c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(text_x, text_y, str(row['Gift Box']))
        
        gift_note_x = gift_box_x + frame_width + 0.2 * inch
        c.rect(gift_note_x, frame_y, frame_width, frame_height, stroke=1, fill=0)
        
        checkbox_x = gift_note_x + (frame_width - checkbox_size) / 2
        is_gift_note_checked = (row['Gift Note'] == "YES")
        draw_checkbox(c, checkbox_x, checkbox_y, checkbox_size, is_gift_note_checked)
        
        text_x = gift_note_x + frame_width / 2
        text_y = frame_y + frame_height - 0.60 * inch
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(text_x, text_y, "GIFT NOTE")
        
        text_y -= 0.25 * inch
        if row['Gift Note'] == "YES":
            c.setFont("Helvetica-BoldOblique", 14)
        else:
            c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(text_x, text_y, str(row['Gift Note']))

        c.showPage()

    c.save()
    buf.seek(0)
    return buf

def generate_gift_message_labels(dataframe):
    buf = BytesIO()
    page_size = landscape((4 * inch, 6 * inch))
    c = canvas.Canvas(buf, pagesize=page_size)
    W, H = page_size
    gift_orders = dataframe[dataframe['Gift Message'] != ""]

    if len(gift_orders) == 0:
        c.setFont("Helvetica", 14)
        c.drawCentredString(W / 2, H / 2, "No gift messages found in orders")
        c.showPage()
    else:
        for _, row in gift_orders.iterrows():
            c.setStrokeColor(colors.black)
            c.setLineWidth(3)
            c.rect(0.4 * inch, 0.4 * inch, W - 0.8 * inch, H - 0.8 * inch, stroke=1, fill=0)
            c.setFont("Times-BoldItalic", 18)
            message = row['Gift Message']
            words = message.split()
            lines = []
            current_line = []
            max_width = W - 1.2 * inch
            for word in words:
                test_line = ' '.join(current_line + [word])
                if c.stringWidth(test_line, "Times-BoldItalic", 18) < max_width:
                    current_line.append(word)
                else:
                    if current_line: lines.append(' '.join(current_line))
                    current_line = [word]
            if current_line: lines.append(' '.join(current_line))
            total_height = len(lines) * 0.3 * inch
            y = (H + total_height) / 2
            for line in lines:
                c.drawCentredString(W / 2, y, line)
                y -= 0.3 * inch
            c.showPage()
    c.save()
    buf.seek(0)
    return buf

def generate_summary_pdf(dataframe, summary_stats):
    buf = BytesIO()
    page_size = A4
    c = canvas.Canvas(buf, pagesize=page_size)
    W, H = page_size
    left = 0.75 * inch
    right = W - 0.75 * inch
    top = H - 0.75 * inch
    y = top
    
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(W / 2, y, "END OF DAY SUMMARY")
    y -= 0.3 * inch
    from datetime import datetime
    c.setFont("Helvetica", 14)
    c.drawCentredString(W / 2, y, f"Report Date: {datetime.now().strftime('%B %d, %Y')}")
    y -= 0.5 * inch
    
    c.setStrokeColor(colors.black)
    c.setLineWidth(2)
    box_height = 2.5 * inch
    box_y = y - box_height
    c.rect(left, box_y, right - left, box_height, stroke=1, fill=0)
    
    y -= 0.3 * inch
    c.setFont("Helvetica-Bold", 16)
    col1_x = left + 0.5 * inch
    col2_x = W / 2 + 0.5 * inch
    
    c.drawString(col1_x, y, "Total Blankets:")
    c.drawRightString(col2_x - 0.3 * inch, y, str(summary_stats['total_blankets']))
    c.drawString(col2_x, y, "Total Beanies:")
    c.drawRightString(right - 0.5 * inch, y, str(summary_stats['total_beanies']))
    y -= 0.35 * inch
    c.drawString(col1_x, y, "Total Orders:")
    c.drawRightString(col2_x - 0.3 * inch, y, str(summary_stats['total_orders']))
    c.drawString(col2_x, y, "Gift Boxes:")
    c.drawRightString(right - 0.5 * inch, y, str(summary_stats['gift_boxes']))
    y -= 0.35 * inch
    c.drawString(col1_x, y, "Blanket Only:")
    c.drawRightString(col2_x - 0.3 * inch, y, str(summary_stats['blanket_only']))
    c.drawString(col2_x, y, "Gift Messages:")
    c.drawRightString(right - 0.5 * inch, y, str(summary_stats['gift_messages']))
    
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# --------------------------------------
# CORE LOGIC: Merge & QC Dashboard
# --------------------------------------
def merge_shipping_and_manufacturing_labels(shipping_pdf_bytes, manufacturing_pdf_bytes, order_dataframe):
    try:
        # 1. Index Manufacturing Labels
        mfg_reader = PdfReader(manufacturing_pdf_bytes)
        mfg_map = {} 
        current_mfg_page_idx = 0
        
        for _, row in order_dataframe.iterrows():
            raw_name = str(row['Buyer Name']).strip().upper()
            raw_name = " ".join(raw_name.split()) # Clean spaces
            
            if raw_name not in mfg_map:
                mfg_map[raw_name] = []
            
            if current_mfg_page_idx < len(mfg_reader.pages):
                mfg_map[raw_name].append(mfg_reader.pages[current_mfg_page_idx])
                current_mfg_page_idx += 1

        known_buyers = list(mfg_map.keys())
        qc_tracker = {name: "❌ MISSING" for name in known_buyers}

        # 2. Process Shipping Labels
        output_pdf = PdfWriter()
        shipping_pdf_bytes.seek(0)
        processed_count = 0
        matched_count = 0
        
        with pdfplumber.open(shipping_pdf_bytes) as plist:
            ship_reader = PdfReader(shipping_pdf_bytes)
            
            for i, page in enumerate(plist.pages):
                text = page.extract_text() or ""
                text = text.upper()
                found_name = None
                
                # Strategy A: Look for "SHIP TO"
                ship_to_match = re.search(r"SHIP\s*TO:?\s*\n+([^\n]+)", text)
                if ship_to_match:
                    candidate = ship_to_match.group(1).strip()
                    matches = get_close_matches(candidate, known_buyers, n=1, cutoff=0.8)
                    if matches: found_name = matches[0]

                # Strategy B: Scan full text
                if not found_name:
                    for buyer in known_buyers:
                        if buyer in text:
                            found_name = buyer
                            break
                
                # Strategy C: OCR Fallback
                if not found_name and len(text) < 50: 
                    try:
                        images = convert_from_bytes(shipping_pdf_bytes.getvalue(), first_page=i+1, last_page=i+1, dpi=150)
                        if images:
                            ocr_text = pytesseract.image_to_string(images[0]).upper()
                            for buyer in known_buyers:
                                if buyer in ocr_text:
                                    found_name = buyer
                                    break
                    except: pass

                # Construct PDF
                output_pdf.add_page(ship_reader.pages[i])
                processed_count += 1
                
                if found_name and found_name in mfg_map:
                    pages_to_add = mfg_map[found_name]
                    for p in pages_to_add:
                        output_pdf.add_page(p)
                        matched_count += 1
                    qc_tracker[found_name] = f"✅ MATCHED (Page {i+1})"
                    del mfg_map[found_name]

        # 3. Handle Orphans
        if len(mfg_map) > 0:
            for buyer, pages in mfg_map.items():
                for p in pages:
                    output_pdf.add_page(p)

        output_buffer = BytesIO()
        output_pdf.write(output_buffer)
        output_buffer.seek(0)
        
        # Generate QC Dataframe
        qc_data = [{"Buyer Name": name, "Status": status} for name, status in qc_tracker.items()]
        qc_df = pd.DataFrame(qc_data)
        qc_df = qc_df.sort_values(by="Status", ascending=False)
        
        return output_buffer, processed_count, matched_count, qc_df

    except Exception as e:
        st.error(f"Merge Error: {str(e)}")
        return BytesIO(), 0, 0, pd.DataFrame()

# --------------------------------------
# MAIN APP INTERFACE
# --------------------------------------
with st.sidebar:
    st.markdown("# 🧵 Blanket Manager")
    st.markdown("### Version 10.5 (Fixed & Full)")
    st.markdown("---")
    st.markdown('<a href="#upload-order" class="nav-link">📄 Upload Order</a>', unsafe_allow_html=True)
    st.markdown('<a href="#dashboard" class="nav-link">📊 Dashboard</a>', unsafe_allow_html=True)
    st.markdown('<a href="#color-analytics" class="nav-link">🎨 Color Analytics</a>', unsafe_allow_html=True)
    st.markdown('<a href="#bobbin-setup" class="nav-link">🧵 Bobbin Setup</a>', unsafe_allow_html=True)
    st.markdown('<a href="#generate-labels" class="nav-link">📥 Generate Labels</a>', unsafe_allow_html=True)
    st.markdown('<a href="#label-merge" class="nav-link">🔄 Label Merge</a>', unsafe_allow_html=True)
    st.markdown('<a href="#airtable-sync" class="nav-link">☁️ Airtable Sync</a>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown('<div class="status-indicator"><div class="status-dot"></div><span>System Ready</span></div>', unsafe_allow_html=True)

st.title("🧵 Amazon Blanket Order Manager")
st.markdown("**Professional order processing & label generation system**")
st.markdown("---")

# 1. Upload
st.markdown('<a id="upload-order"></a>', unsafe_allow_html=True)
st.markdown("## 📄 Upload Order")
uploaded = st.file_uploader("Drop your Amazon packing slip PDF here", type=["pdf"])

if uploaded:
    records = []
    with pdfplumber.open(uploaded) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            oid = re.search(r"Order ID:\s*([\d\-]+)", text)
            odate = re.search(r"Order Date:\s*([A-Za-z]{3,},?\s*[A-Za-z]+\s*\d{1,2},?\s*\d{4})", text)
            buyer = re.search(r"Ship To:\s*([\s\S]*?)Order ID:", text)
            
            blocks = re.split(r"(?=Customizations:)", text)
            for block in blocks:
                if "Customizations:" not in block: continue
                qty = re.search(r"Quantity\s*\n\s*(\d+)", block)
                quantity = qty.group(1) if qty else "1"
                color = re.search(r"Color:\s*([^\n]+)", block)
                thread = re.search(r"Thread Color:\s*([^\n]+)", block, re.IGNORECASE)
                name = re.search(r"Name:\s*([^\n]+)", block)
                gift_msg = re.search(r"Gift Message:\s*([\s\S]*?)(?=\n(?:Grand total|Returning|Visit|Quantity|$))", block, re.IGNORECASE)
                
                records.append({
                    "Order ID": oid.group(1) if oid else "",
                    "Order Date": odate.group(1) if odate else "",
                    "Buyer Name": buyer.group(1).strip().split('\n')[0] if buyer else "Unknown",
                    "Quantity": quantity,
                    "Blanket Color": clean_text(color.group(1)) if color else "",
                    "Thread Color": translate_thread_color(clean_text(thread.group(1))) if thread else "",
                    "Customization Name": clean_text(name.group(1)) if name else "",
                    "Include Beanie": "YES" if re.search(r"Beanie:\s*Yes", block, re.IGNORECASE) else "NO",
                    "Gift Box": "YES" if re.search(r"Gift Box.*Yes", block, re.IGNORECASE) else "NO",
                    "Gift Note": "YES" if re.search(r"Gift Message:", block, re.IGNORECASE) else "NO",
                    "Gift Message": clean_text(gift_msg.group(1)) if gift_msg else ""
                })

    df = pd.DataFrame(records)
    df.index = df.index + 1
    
    if not df.empty:
        st.success(f"✅ Parsed {len(df)} items from {df['Order ID'].nunique()} orders")
        with st.expander("📊 View Order Data"):
            st.dataframe(df, use_container_width=True)
        
        # Stats
        df['Quantity_Int'] = df['Quantity'].astype(int)
        total_blankets = df['Quantity_Int'].sum()
        total_beanies = df[df['Include Beanie'] == 'YES']['Quantity_Int'].sum()
        blanket_counts = df.groupby('Blanket Color')['Quantity_Int'].sum().sort_values(ascending=False)
        thread_counts = df.groupby('Thread Color')['Quantity_Int'].sum().sort_values(ascending=False)
        
        # Dashboard
        st.markdown('<a id="dashboard"></a>', unsafe_allow_html=True)
        st.markdown("## 📊 Order Dashboard")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("Total Blankets", total_blankets)
        with col2: st.metric("Orders", df['Order ID'].nunique())
        with col3: st.metric("Beanies", total_beanies)
        with col4: st.metric("Gift Boxes", len(df[df['Gift Box'] == 'YES']))
        
        st.markdown("---")
        st.markdown('<a id="color-analytics"></a>', unsafe_allow_html=True)
        st.markdown("## 🎨 Color Analytics")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### 🧶 Blanket Colors")
            for c, n in blanket_counts.items(): st.markdown(f"**{c}:** {n}")
        with c2:
            st.markdown("### 🧵 Thread Colors")
            for c, n in thread_counts.items(): st.markdown(f"**{c}:** {n}")

        # Generate Labels
        st.markdown("---")
        st.markdown('<a id="generate-labels"></a>', unsafe_allow_html=True)
        st.markdown("## 📥 Generate & Download")
        
        if 'manufacturing_labels_buffer' not in st.session_state:
            st.session_state.manufacturing_labels_buffer = None
        
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("📦 Manufacturing Labels", use_container_width=True):
                pdf = generate_manufacturing_labels(df)
                st.session_state.manufacturing_labels_buffer = pdf
                st.success("Generated!")
            if st.session_state.manufacturing_labels_buffer:
                st.download_button("⬇️ Download PDF", st.session_state.manufacturing_labels_buffer, "Manufacturing_Labels.pdf", "application/pdf", use_container_width=True)
        
        with c2:
            if st.button("💌 Gift Messages", use_container_width=True):
                pdf = generate_gift_message_labels(df)
                st.session_state.gift_pdf = pdf
                st.success("Generated!")
            if 'gift_pdf' in st.session_state:
                st.download_button("⬇️ Download PDF", st.session_state.gift_pdf, "Gift_Messages.pdf", "application/pdf", use_container_width=True)

        with c3:
            if st.button("📊 Summary Report", use_container_width=True):
                # Simplified summary dict
                summ = {'total_blankets': total_blankets, 'total_beanies': total_beanies, 
                        'total_orders': df['Order ID'].nunique(), 'blanket_only': len(df[df['Include Beanie']=='NO']),
                        'with_beanie': len(df[df['Include Beanie']=='YES']), 'gift_boxes': len(df[df['Gift Box']=='YES']),
                        'gift_messages': len(df[df['Gift Note']=='YES']), 'unique_colors': len(blanket_counts),
                        'blanket_colors': blanket_counts.to_dict(), 'thread_colors': thread_counts.to_dict(),
                        'black_bobbin_total': 0, 'white_bobbin_total': 0, 'black_bobbin_threads': {}, 'white_bobbin_threads': {}}
                pdf = generate_summary_pdf(df, summ)
                st.session_state.sum_pdf = pdf
                st.success("Generated!")
            if 'sum_pdf' in st.session_state:
                st.download_button("⬇️ Download PDF", st.session_state.sum_pdf, "Summary.pdf", "application/pdf", use_container_width=True)

        # Merge
        st.markdown("---")
        st.markdown('<a id="label-merge"></a>', unsafe_allow_html=True)
        st.markdown("## 🔄 Merge Shipping & Manufacturing Labels")
        ship_upload = st.file_uploader("Upload Shipping Labels (PDF)", type=["pdf"], key="ship")
        
        if ship_upload and st.session_state.manufacturing_labels_buffer:
            if st.button("🔀 Merge & Run QC Check", type="primary", use_container_width=True):
                with st.spinner("Merging..."):
                    ship_upload.seek(0)
                    st.session_state.manufacturing_labels_buffer.seek(0)
                    merged, n_ship, n_mfg, qc_df = merge_shipping_and_manufacturing_labels(ship_upload, st.session_state.manufacturing_labels_buffer, df)
                    
                    st.markdown("### ✅ QC Dashboard")
                    missing = len(qc_df[qc_df['Status'].str.contains("MISSING")])
                    if missing > 0: st.error(f"⚠️ {missing} Orders Missing Labels!")
                    else: st.success("🎉 All Matched!")
                    
                    st.dataframe(qc_df, use_container_width=True)
                    st.download_button("⬇️ Download Final Merged PDF", merged, "Final_Merged.pdf", "application/pdf", use_container_width=True)
        
        # Airtable
        st.markdown("---")
        st.markdown('<a id="airtable-sync"></a>', unsafe_allow_html=True)
        st.markdown("## ☁️ Airtable Integration")
        if st.button("🚀 Upload to Airtable", use_container_width=True):
            with st.spinner("Uploading..."):
                c_orders, c_items, errs = upload_to_airtable(df)
            if errs: st.error(f"Errors: {len(errs)}")
            else: st.success(f"Uploaded {c_orders} orders!")
