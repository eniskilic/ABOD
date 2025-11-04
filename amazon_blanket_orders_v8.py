import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch, landscape
from reportlab.lib import colors
import requests
from pypdf import PdfReader, PdfWriter

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
# Custom CSS Styling
# --------------------------------------
st.markdown("""
<style>
    /* Main App Styling */
    .main {
        background: #f5f7fa;
    }
    
    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #2c3e50 0%, #34495e 100%);
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: white;
    }
    
    /* Metric Cards */
    [data-testid="stMetric"] {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border-left: 4px solid #3498db;
    }
    
    [data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    
    [data-testid="stMetric"] label {
        font-size: 0.9em !important;
        color: #7f8c8d !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 2.5em !important;
        font-weight: 700 !important;
        color: #2c3e50 !important;
    }
    
    /* Headers */
    h1 {
        color: #2c3e50;
        font-weight: 700;
        padding-bottom: 10px;
        border-bottom: 3px solid #3498db;
        margin-bottom: 30px;
    }
    
    h2 {
        color: #2c3e50;
        font-weight: 600;
        margin-top: 40px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    h3 {
        color: #34495e;
        font-weight: 600;
        margin-bottom: 15px;
    }
    
    /* Buttons */
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.4);
    }
    
    /* File Uploader */
    [data-testid="stFileUploader"] {
        background: white;
        padding: 30px;
        border-radius: 12px;
        border: 2px dashed #3498db;
        text-align: center;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #2980b9;
        background: #f8f9fa;
    }
    
    /* Info boxes */
    .stAlert {
        border-radius: 10px;
        border-left: 4px solid #3498db;
    }
    
    /* Dataframe */
    [data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    /* Expander */
    [data-testid="stExpander"] {
        background: white;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        margin-bottom: 10px;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Download button */
    .stDownloadButton button {
        background: #27ae60;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
        width: 100%;
    }
    
    .stDownloadButton button:hover {
        background: #229954;
        transform: translateY(-1px);
    }
    
    /* Custom card styling */
    .metric-card {
        background: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border-left: 4px solid #3498db;
        margin-bottom: 20px;
    }
    
    .bobbin-card {
        background: white;
        padding: 25px;
        border-radius: 12px;
        border: 2px solid #e0e0e0;
        margin-bottom: 15px;
    }
    
    .bobbin-card.black {
        border-left: 5px solid #2c3e50;
    }
    
    .bobbin-card.white {
        border-left: 5px solid #95a5a6;
    }
    
    /* Section divider */
    hr {
        border: none;
        border-top: 2px solid #ecf0f1;
        margin: 40px 0;
    }
</style>
""", unsafe_allow_html=True)

# --------------------------------------
# Airtable Configuration
# --------------------------------------
AIRTABLE_PAT = "patD9n2LOJRthfGan.b420b57e48143665f27870484e882266bcfd184fa7c96067fbb1ef8c41424fae"
BASE_ID = "appxoNC3r5NSsTP3U"
ORDERS_TABLE = "Orders"
LINE_ITEMS_TABLE = "Order Line Items"

# --------------------------------------
# Color dictionary for Spanish translation
# --------------------------------------
COLOR_TRANSLATIONS = {
    "white": "Blanco",
    "black": "Negro",
    "brown": "Marrón",
    "blue": "Azul",
    "navy": "Azul Marino",
    "red": "Rojo",
    "pink": "Rosa",
    "light pink": "Rosa Claro",
    "hot pink": "Rosa Fucsia",
    "salmon pink": "Rosa Salmón",
    "purple": "Morado",
    "lilac": "Lila",
    "gray": "Gris",
    "grey": "Gris",
    "gold": "Dorado",
    "silver": "Plateado",
    "beige": "Beige",
    "green": "Verde",
    "olive": "Verde Oliva",
    "yellow": "Amarillo",
    "champagne": "Champán"
}

# --------------------------------------
# Helper Functions
# --------------------------------------
def clean_text(s: str) -> str:
    """Cleans unwanted symbols and color codes."""
    if not s:
        return ""
    s = re.sub(r"\(#?[A-Fa-f0-9]{3,6}\)", "", s)
    s = re.sub(r"■|Seller Name|Your Orders|Returning your item:", "", s)
    s = re.sub(r"\(Most popular\)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def translate_thread_color(color):
    """Adds Spanish translation."""
    if not color:
        return color
    base = color.strip()
    for eng, esp in COLOR_TRANSLATIONS.items():
        if eng.lower() in base.lower():
            return f"{base} ({esp})"
    return base

def get_bobbin_color(thread_color):
    """Determine bobbin color based on thread color"""
    thread_lower = thread_color.lower()
    if 'navy' in thread_lower or 'black' in thread_lower or 'negro' in thread_lower:
        return 'Black Bobbin'
    else:
        return 'White Bobbin'

def draw_checkbox(canvas_obj, x, y, size, is_checked):
    """Draw a checkbox at position (x, y) with given size."""
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
# FIXED: Label Merging Function
# --------------------------------------
def merge_shipping_and_manufacturing_labels(shipping_pdf_bytes, manufacturing_pdf_bytes, order_dataframe):
    """
    Merge shipping labels with manufacturing labels.
    Handles orders with multiple items (one shipping label, multiple manufacturing labels).
    """
    try:
        shipping_pdf = PdfReader(shipping_pdf_bytes)
        manufacturing_pdf = PdfReader(manufacturing_pdf_bytes)
        
        # CRITICAL FIX: Preserve the order of orders as they appear in the dataframe
        seen_orders = []
        order_item_counts = []
        
        for order_id in order_dataframe['Order ID']:
            if order_id not in seen_orders:
                seen_orders.append(order_id)
                item_count = len(order_dataframe[order_dataframe['Order ID'] == order_id])
                order_item_counts.append(item_count)
        
        # Build mapping: shipping label index -> list of manufacturing label indices
        shipping_to_mfg = {}
        mfg_index = 0
        
        for shipping_index, item_count in enumerate(order_item_counts):
            shipping_to_mfg[shipping_index] = list(range(mfg_index, mfg_index + item_count))
            mfg_index += item_count
        
        # Create merged PDF
        output_pdf = PdfWriter()
        total_shipping_labels = len(shipping_to_mfg)
        
        for ship_idx in range(total_shipping_labels):
            if ship_idx >= len(shipping_pdf.pages):
                break
                
            output_pdf.add_page(shipping_pdf.pages[ship_idx])
            
            if ship_idx in shipping_to_mfg:
                for mfg_idx in shipping_to_mfg[ship_idx]:
                    if mfg_idx < len(manufacturing_pdf.pages):
                        output_pdf.add_page(manufacturing_pdf.pages[mfg_idx])
        
        output_buffer = BytesIO()
        output_pdf.write(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer, len(shipping_to_mfg), sum(len(v) for v in shipping_to_mfg.values())
        
    except Exception as e:
        st.error(f"Error merging labels: {str(e)}")
        return None, 0, 0

# --------------------------------------
# Airtable Functions
# --------------------------------------
def get_existing_order_ids():
    """Fetch all existing Order IDs from Airtable"""
    headers = {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json"
    }
    
    existing_orders = set()
    offset = None
    
    try:
        while True:
            url = f"https://api.airtable.com/v0/{BASE_ID}/{ORDERS_TABLE}"
            params = {"fields[]": "Order ID"}
            
            if offset:
                params["offset"] = offset
            
            response = requests.get(url, headers=headers, params=params)
            
            if response.status_code == 200:
                data = response.json()
                records = data.get("records", [])
                
                for record in records:
                    order_id = record.get("fields", {}).get("Order ID")
                    if order_id:
                        existing_orders.add(order_id)
                
                offset = data.get("offset")
                if not offset:
                    break
            else:
                st.warning(f"Could not fetch existing orders: {response.text}")
                break
                
    except Exception as e:
        st.warning(f"Error checking for duplicates: {str(e)}")
    
    return existing_orders

def upload_to_airtable(dataframe):
    """Upload parsed orders to Airtable with duplicate detection"""
    headers = {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json"
    }
    
    st.info("🔍 Checking for duplicate orders...")
    existing_order_ids = get_existing_order_ids()
    
    unique_orders = dataframe[['Order ID', 'Order Date', 'Buyer Name']].drop_duplicates(subset=['Order ID'])
    
    new_orders = unique_orders[~unique_orders['Order ID'].isin(existing_order_ids)]
    duplicate_orders = unique_orders[unique_orders['Order ID'].isin(existing_order_ids)]
    
    if len(duplicate_orders) > 0:
        st.warning(f"⚠️ Found {len(duplicate_orders)} duplicate order(s) that already exist in Airtable")
        with st.expander("View Duplicate Orders (will be skipped)"):
            for _, dup in duplicate_orders.iterrows():
                st.write(f"• {dup['Order ID']} - {dup['Buyer Name']}")
    
    if len(new_orders) == 0:
        st.info("ℹ️ All orders already exist in Airtable. Nothing to upload.")
        return 0, 0, []
    
    st.success(f"✅ Found {len(new_orders)} new order(s) to upload")
    
    orders_created = 0
    line_items_created = 0
    errors = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_orders = len(new_orders)
    
    for idx, (_, order_row) in enumerate(new_orders.iterrows()):
        order_id = order_row['Order ID']
        
        try:
            status_text.text(f"Creating order {idx + 1}/{total_orders}: {order_id}")
            
            order_payload = {
                "records": [{
                    "fields": {
                        "Order ID": order_id,
                        "Order Date": order_row['Order Date'],
                        "Buyer Name": order_row['Buyer Name'],
                        "Status": "New"
                    }
                }]
            }
            
            response = requests.post(
                f"https://api.airtable.com/v0/{BASE_ID}/{ORDERS_TABLE}",
                headers=headers,
                json=order_payload
            )
            
            if response.status_code == 200:
                airtable_order_id = response.json()["records"][0]["id"]
                orders_created += 1
                
                order_items = dataframe[dataframe['Order ID'] == order_id]
                
                for _, item_row in order_items.iterrows():
                    line_item_payload = {
                        "records": [{
                            "fields": {
                                "Buyer Name": item_row['Buyer Name'],
                                "Customization Name": item_row['Customization Name'],
                                "Order ID": [airtable_order_id],
                                "Quantity": int(item_row['Quantity']),
                                "Blanket Color": item_row['Blanket Color'],
                                "Thread Color": item_row['Thread Color'],
                                "Include Beanie": item_row['Include Beanie'],
                                "Gift Box": item_row['Gift Box'],
                                "Gift Note": item_row['Gift Note'],
                                "Gift Message": item_row['Gift Message'],
                                "Bobbin Color": get_bobbin_color(item_row['Thread Color']),
                                "Status": "Pending"
                            }
                        }]
                    }
                    
                    item_response = requests.post(
                        f"https://api.airtable.com/v0/{BASE_ID}/{LINE_ITEMS_TABLE}",
                        headers=headers,
                        json=line_item_payload
                    )
                    
                    if item_response.status_code == 200:
                        line_items_created += 1
                    else:
                        errors.append(f"Error creating line item for {order_id}: {item_response.text}")
            else:
                errors.append(f"Error creating order {order_id}: {response.text}")
        
        except Exception as e:
            errors.append(f"Exception for order {order_id}: {str(e)}")
        
        progress_bar.progress((idx + 1) / total_orders)
    
    status_text.empty()
    progress_bar.empty()
    
    return orders_created, line_items_created, errors

# --------------------------------------
# PDF Generation Functions
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
        c.drawString(left + 0.1 * inch, text_y, f"BLANKET COLOR: {row['Blanket Color'].upper()}")
        
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
        
        # BEANIE
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
        c.drawCentredString(text_x, text_y, row['Include Beanie'])
        
        # GIFT BOX
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
        c.drawCentredString(text_x, text_y, row['Gift Box'])
        
        # GIFT NOTE
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
        c.drawCentredString(text_x, text_y, row['Gift Note'])

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
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))

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
    from reportlab.lib.pagesizes import A4
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
    today = datetime.now().strftime("%B %d, %Y")
    c.setFont("Helvetica", 14)
    c.drawCentredString(W / 2, y, f"Report Date: {today}")
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
    y -= 0.35 * inch
    
    c.drawString(col1_x, y, "With Beanie:")
    c.drawRightString(col2_x - 0.3 * inch, y, str(summary_stats['with_beanie']))
    c.drawString(col2_x, y, "Unique Colors:")
    c.drawRightString(right - 0.5 * inch, y, str(summary_stats['unique_colors']))
    
    y = box_y - 0.5 * inch
    
    c.setFont("Helvetica-Bold", 18)
    c.drawString(left, y, "Blanket Color Breakdown")
    y -= 0.3 * inch
    
    c.setStrokeColor(colors.grey)
    c.setLineWidth(1)
    c.line(left, y, right, y)
    y -= 0.25 * inch
    
    c.setFont("Helvetica", 14)
    for color, count in summary_stats['blanket_colors'].items():
        if y < 2 * inch:
            c.showPage()
            y = top
            c.setFont("Helvetica", 14)
        
        c.drawString(left + 0.3 * inch, y, f"{color}:")
        c.drawRightString(right - 0.3 * inch, y, str(count))
        y -= 0.22 * inch
    
    y -= 0.3 * inch
    
    if y < 3 * inch:
        c.showPage()
        y = top
    
    c.setFont("Helvetica-Bold", 18)
    c.drawString(left, y, "Thread Color Breakdown")
    y -= 0.3 * inch
    
    c.setStrokeColor(colors.grey)
    c.setLineWidth(1)
    c.line(left, y, right, y)
    y -= 0.25 * inch
    
    c.setFont("Helvetica", 14)
    for color, count in summary_stats['thread_colors'].items():
        if y < 1.5 * inch:
            c.showPage()
            y = top
            c.setFont("Helvetica", 14)
        
        c.drawString(left + 0.3 * inch, y, f"{color}:")
        c.drawRightString(right - 0.3 * inch, y, str(count))
        y -= 0.22 * inch
    
    y -= 0.5 * inch
    
    if y < 4 * inch:
        c.showPage()
        y = top
    
    c.setFont("Helvetica-Bold", 18)
    c.drawString(left, y, "Bobbin Color Setup")
    y -= 0.3 * inch
    
    c.setStrokeColor(colors.grey)
    c.setLineWidth(1)
    c.line(left, y, right, y)
    y -= 0.3 * inch
    
    c.setFont("Helvetica-Bold", 16)
    c.drawString(left + 0.3 * inch, y, "⚫ Black Bobbin")
    c.drawRightString(right - 0.3 * inch, y, f"Total: {summary_stats['black_bobbin_total']}")
    y -= 0.25 * inch
    
    c.setFont("Helvetica", 13)
    for color, count in summary_stats['black_bobbin_threads'].items():
        if y < 1.5 * inch:
            c.showPage()
            y = top
            c.setFont("Helvetica", 13)
        c.drawString(left + 0.6 * inch, y, f"• {color}:")
        c.drawRightString(right - 0.3 * inch, y, str(count))
        y -= 0.2 * inch
    
    y -= 0.25 * inch
    
    if y < 2 * inch:
        c.showPage()
        y = top
    
    c.setFont("Helvetica-Bold", 16)
    c.drawString(left + 0.3 * inch, y, "⚪ White Bobbin")
    c.drawRightString(right - 0.3 * inch, y, f"Total: {summary_stats['white_bobbin_total']}")
    y -= 0.25 * inch
    
    c.setFont("Helvetica", 13)
    for color, count in summary_stats['white_bobbin_threads'].items():
        if y < 1.5 * inch:
            c.showPage()
            y = top
            c.setFont("Helvetica", 13)
        c.drawString(left + 0.6 * inch, y, f"• {color}:")
        c.drawRightString(right - 0.3 * inch, y, str(count))
        y -= 0.2 * inch
    
    c.save()
    buf.seek(0)
    return buf

# --------------------------------------
# SIDEBAR
# --------------------------------------
with st.sidebar:
    st.markdown("# 🧵 Blanket Manager")
    st.markdown("### Version 10.1")
    st.markdown("---")
    
    st.markdown("#### 📋 Navigation")
    st.markdown("• Dashboard (current)")
    st.markdown("• Manufacturing")
    st.markdown("• Label Merge")
    st.markdown("• Airtable Sync")
    st.markdown("• Reports")
    
    st.markdown("---")
    
    st.markdown("#### ✨ Features")
    st.markdown("✓ PDF Parsing")
    st.markdown("✓ Label Generation")
    st.markdown("✓ Order Merging")
    st.markdown("✓ Cloud Sync")
    st.markdown("✓ Spanish Translation")
    st.markdown("✓ Duplicate Detection")
    
    st.markdown("---")
    st.markdown("**Status:** 🟢 System Ready")

# --------------------------------------
# MAIN CONTENT
# --------------------------------------
st.title("🧵 Amazon Blanket Order Manager")

st.markdown("""
**Professional order processing & label generation system**  
Parse Amazon PDFs • Generate labels • Merge shipments • Sync to Airtable
""")

st.markdown("---")

# File Upload Section
st.markdown("## 📄 Upload Order")
uploaded = st.file_uploader(
    "Drop your Amazon packing slip PDF here",
    type=["pdf"],
    help="Upload the packing slip PDF from your Amazon orders"
)

# --------------------------------------
# Parse PDF
# --------------------------------------
if uploaded:
    st.info("⏳ Reading and parsing your PDF...")

    all_pages = []
    with pdfplumber.open(uploaded) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_pages.append(text)

    records = []

    for page_text in all_pages:
        buyer_match = re.search(r"Ship To:\s*([\s\S]*?)Order ID:", page_text)
        buyer_name = ""
        if buyer_match:
            lines = [l.strip() for l in buyer_match.group(1).splitlines() if l.strip()]
            if lines:
                buyer_name = lines[0]

        order_id = ""
        order_date = ""
        m_id = re.search(r"Order ID:\s*([\d\-]+)", page_text)
        if m_id:
            order_id = m_id.group(1).strip()
        m_date = re.search(r"Order Date:\s*([A-Za-z]{3,},?\s*[A-Za-z]+\s*\d{1,2},?\s*\d{4})", page_text)
        if m_date:
            order_date = m_date.group(1).strip()

        blocks = re.split(r"(?=Customizations:)", page_text)
        for block in blocks:
            if "Customizations:" not in block:
                continue

            qty_match = re.search(r"Quantity\s*\n\s*(\d+)", block)
            quantity = qty_match.group(1) if qty_match else "1"

            blanket_color = ""
            thread_color = ""
            b_match = re.search(r"Color:\s*([^\n]+)", block)
            if b_match:
                blanket_color = clean_text(b_match.group(1))
            t_match = re.search(r"Thread Color:\s*([^\n]+)", block, re.IGNORECASE)
            if t_match:
                thread_color = translate_thread_color(clean_text(t_match.group(1)))

            name_match = re.search(r"Name:\s*([^\n]+)", block)
            customization_name = clean_text(name_match.group(1)) if name_match else ""

            beanie = "YES" if re.search(r"Personalized Baby Beanie:\s*Yes", block, re.IGNORECASE) else "NO"
            gift_box = "YES" if re.search(r"Gift Box\s*&\s*Gift Card:\s*Yes", block, re.IGNORECASE) else "NO"
            gift_note = "YES" if re.search(r"Gift Message:", block, re.IGNORECASE) else "NO"

            gift_msg_match = re.search(
                r"Gift Message:\s*([\s\S]*?)(?=\n(?:Grand total|Returning your item|Visit|Quantity|Order Totals|$))",
                block,
                re.IGNORECASE
            )
            gift_message = clean_text(gift_msg_match.group(1)) if gift_msg_match else ""

            records.append({
                "Order ID": order_id,
                "Order Date": order_date,
                "Buyer Name": buyer_name,
                "Quantity": quantity,
                "Blanket Color": blanket_color,
                "Thread Color": thread_color,
                "Customization Name": customization_name,
                "Include Beanie": beanie,
                "Gift Box": gift_box,
                "Gift Note": gift_note,
                "Gift Message": gift_message
            })

    if not records:
        st.error("❌ No orders detected. Please check your PDF format.")
        st.stop()

    df = pd.DataFrame(records)
    df.index = df.index + 1
    
    st.success(f"✅ Successfully parsed {len(df)} line items from {df['Order ID'].nunique()} orders")
    
    with st.expander("📊 View Order Data"):
        st.dataframe(df, use_container_width=True)

    # --------------------------------------
    # Calculate Summary Statistics
    # --------------------------------------
    df['Quantity_Int'] = df['Quantity'].astype(int)
    
    total_blankets = df['Quantity_Int'].sum()
    total_beanies = df[df['Include Beanie'] == 'YES']['Quantity_Int'].sum()
    total_orders = df['Order ID'].nunique()
    orders_blanket_only = len(df[df['Include Beanie'] == 'NO'])
    orders_with_beanie = len(df[df['Include Beanie'] == 'YES'])
    gift_boxes_needed = len(df[df['Gift Box'] == 'YES'])
    gift_messages_needed = len(df[df['Gift Note'] == 'YES'])
    
    blanket_color_counts = df.groupby('Blanket Color')['Quantity_Int'].sum().sort_values(ascending=False)
    thread_color_counts = df.groupby('Thread Color')['Quantity_Int'].sum().sort_values(ascending=False)
    
    df['Bobbin_Color'] = df['Thread Color'].apply(get_bobbin_color)
    bobbin_counts = df.groupby('Bobbin_Color')['Quantity_Int'].sum()
    
    black_bobbin_df = df[df['Bobbin_Color'] == 'Black Bobbin']
    white_bobbin_df = df[df['Bobbin_Color'] == 'White Bobbin']
    
    black_bobbin_threads = black_bobbin_df.groupby('Thread Color')['Quantity_Int'].sum().sort_values(ascending=False)
    white_bobbin_threads = white_bobbin_df.groupby('Thread Color')['Quantity_Int'].sum().sort_values(ascending=False)

    # --------------------------------------
    # Dashboard Metrics
    # --------------------------------------
    st.markdown("---")
    st.markdown("## 📊 Order Dashboard")
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric("Total Blankets", total_blankets)
    with col2:
        st.metric("Total Orders", total_orders)
    with col3:
        st.metric("Beanies", total_beanies)
    with col4:
        st.metric("Gift Boxes", gift_boxes_needed)
    with col5:
        st.metric("Gift Messages", gift_messages_needed)
    with col6:
        st.metric("Unique Colors", len(blanket_color_counts))
    
    col7, col8 = st.columns(2)
    with col7:
        st.metric("Blanket Only", orders_blanket_only)
    with col8:
        st.metric("With Beanie", orders_with_beanie)
    
    # --------------------------------------
    # Color Breakdown
    # --------------------------------------
    st.markdown("---")
    st.markdown("## 🎨 Color Analytics")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### 🧶 Blanket Colors")
        for color, count in blanket_color_counts.items():
            st.markdown(f"**{color}:** {count}")
    
    with col_right:
        st.markdown("### 🧵 Thread Colors")
        for color, count in thread_color_counts.items():
            st.markdown(f"**{color}:** {count}")
    
    # --------------------------------------
    # Bobbin Setup Section
    # --------------------------------------
    st.markdown("---")
    st.markdown("## 🧵 Bobbin Color Configuration")
    
    col_bobbin1, col_bobbin2 = st.columns(2)
    
    with col_bobbin1:
        st.markdown("### ⚫ Black Bobbin")
        st.metric("Total Items", bobbin_counts.get('Black Bobbin', 0))
        if len(black_bobbin_threads) > 0:
            for color, count in black_bobbin_threads.items():
                st.markdown(f"• **{color}:** {count}")
        else:
            st.markdown("_No items_")
    
    with col_bobbin2:
        st.markdown("### ⚪ White Bobbin")
        st.metric("Total Items", bobbin_counts.get('White Bobbin', 0))
        if len(white_bobbin_threads) > 0:
            for color, count in white_bobbin_threads.items():
                st.markdown(f"• **{color}:** {count}")
        else:
            st.markdown("_No items_")

    # --------------------------------------
    # Generate Labels Section
    # --------------------------------------
    st.markdown("---")
    st.markdown("## 📥 Generate & Download")
    
    # Store manufacturing labels in session state
    if 'manufacturing_labels_buffer' not in st.session_state:
        st.session_state.manufacturing_labels_buffer = None
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📦 Manufacturing Labels", use_container_width=True):
            with st.spinner("Generating manufacturing labels..."):
                pdf_data = generate_manufacturing_labels(df)
                st.session_state.manufacturing_labels_buffer = pdf_data
            st.success("✅ Labels generated!")
            st.download_button(
                label="⬇️ Download Manufacturing Labels",
                data=pdf_data,
                file_name="Manufacturing_Labels.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col2:
        gift_count = len(df[df['Gift Message'] != ""])
        if st.button(f"💌 Gift Messages ({gift_count})", use_container_width=True):
            with st.spinner("Generating gift message labels..."):
                pdf_data = generate_gift_message_labels(df)
            st.success("✅ Labels generated!")
            st.download_button(
                label="⬇️ Download Gift Message Labels",
                data=pdf_data,
                file_name="Gift_Message_Labels.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col3:
        if st.button("📊 Summary Report", use_container_width=True):
            with st.spinner("Generating summary report..."):
                summary_stats = {
                    'total_blankets': total_blankets,
                    'total_beanies': total_beanies,
                    'total_orders': total_orders,
                    'blanket_only': orders_blanket_only,
                    'with_beanie': orders_with_beanie,
                    'gift_boxes': gift_boxes_needed,
                    'gift_messages': gift_messages_needed,
                    'unique_colors': len(blanket_color_counts),
                    'blanket_colors': blanket_color_counts.to_dict(),
                    'thread_colors': thread_color_counts.to_dict(),
                    'black_bobbin_total': int(bobbin_counts.get('Black Bobbin', 0)),
                    'white_bobbin_total': int(bobbin_counts.get('White Bobbin', 0)),
                    'black_bobbin_threads': black_bobbin_threads.to_dict() if len(black_bobbin_threads) > 0 else {},
                    'white_bobbin_threads': white_bobbin_threads.to_dict() if len(white_bobbin_threads) > 0 else {}
                }
                pdf_data = generate_summary_pdf(df, summary_stats)
            st.success("✅ Report generated!")
            st.download_button(
                label="⬇️ Download Summary PDF",
                data=pdf_data,
                file_name="Daily_Summary_Report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    # --------------------------------------
    # Label Merging Section
    # --------------------------------------
    st.markdown("---")
    st.markdown("## 🔄 Merge Shipping & Manufacturing Labels")
    
    st.info("""
    **Instructions for Label Merging:**
    1. Generate Manufacturing Labels above (click the button)
    2. Upload your shipping labels PDF from Amazon/UPS
    3. Click merge to create a combined PDF
    
    ✨ **Version 10.1:** Fixed label ordering for multi-item orders!
    """)
    
    shipping_labels_upload = st.file_uploader(
        "📤 Upload Shipping Labels PDF",
        type=["pdf"],
        key="shipping_labels",
        help="Upload the shipping labels PDF from Amazon or your carrier"
    )
    
    if shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        col_merge1, col_merge2 = st.columns([3, 1])
        
        with col_merge1:
            if st.button("🔀 Merge Labels Now", type="primary", use_container_width=True):
                with st.spinner("Merging shipping and manufacturing labels..."):
                    shipping_labels_upload.seek(0)
                    st.session_state.manufacturing_labels_buffer.seek(0)
                    
                    merged_pdf, num_shipping, num_manufacturing = merge_shipping_and_manufacturing_labels(
                        shipping_labels_upload,
                        st.session_state.manufacturing_labels_buffer,
                        df
                    )
                    
                    if merged_pdf:
                        st.success(f"✅ Successfully merged {num_shipping} shipping labels with {num_manufacturing} manufacturing labels!")
                        
                        multi_item_orders = df.groupby('Order ID').size()
                        multi_item_orders = multi_item_orders[multi_item_orders > 1]
                        
                        if len(multi_item_orders) > 0:
                            with st.expander(f"ℹ️ Found {len(multi_item_orders)} order(s) with multiple items"):
                                for order_id, count in multi_item_orders.items():
                                    buyer = df[df['Order ID'] == order_id]['Buyer Name'].iloc[0]
                                    st.write(f"• {buyer} ({order_id}): {count} blankets")
                        
                        st.download_button(
                            label="⬇️ Download Merged Labels PDF",
                            data=merged_pdf,
                            file_name="Merged_Shipping_Manufacturing_Labels.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
        
        with col_merge2:
            st.metric("Total Orders", df['Order ID'].nunique())
            st.metric("Total Items", len(df))
    
    elif shipping_labels_upload and not st.session_state.manufacturing_labels_buffer:
        st.warning("⚠️ Please generate Manufacturing Labels first (click the button above)")
    
    elif not shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        st.info("📤 Upload your shipping labels PDF above to enable merging")

    # --------------------------------------
    # Airtable Upload Section
    # --------------------------------------
    st.markdown("---")
    st.markdown("## ☁️ Airtable Integration")
    
    st.info("📤 Upload these orders to your Airtable base. Duplicate orders will be automatically detected and skipped.")
    
    if st.button("🚀 Upload to Airtable", type="primary", use_container_width=True):
        with st.spinner("Uploading to Airtable..."):
            orders_created, line_items_created, errors = upload_to_airtable(df)
        
        if errors:
            st.error(f"⚠️ Upload completed with {len(errors)} errors")
            with st.expander("View Errors"):
                for error in errors:
                    st.write(f"• {error}")
        else:
            st.success("✅ Successfully uploaded all orders!")
        
        col_result1, col_result2 = st.columns(2)
        with col_result1:
            st.metric("Orders Created", orders_created)
        with col_result2:
            st.metric("Line Items Created", line_items_created)
        
        st.info("🔗 Go to your Airtable base to view and manage orders!")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #7f8c8d; padding: 20px;'>
    <p><strong>Amazon Blanket Order Manager v10.1</strong></p>
    <p>Professional order processing & label generation system</p>
</div>
""", unsafe_allow_html=True)
