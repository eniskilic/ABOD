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
# Dark Mode Custom CSS Styling
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
    
    /* Alert Boxes */
    .stAlert {
        background-color: #2d3748;
        color: #e4e6eb;
        border-radius: 10px;
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
    "white": "Blanco", "black": "Negro", "brown": "Marrón", "blue": "Azul",
    "navy": "Azul Marino", "red": "Rojo", "pink": "Rosa", "light pink": "Rosa Claro",
    "hot pink": "Rosa Fucsia", "salmon pink": "Rosa Salmón", "purple": "Morado",
    "lilac": "Lila", "gray": "Gris", "grey": "Gris", "gold": "Dorado",
    "silver": "Plateado", "beige": "Beige", "green": "Verde", "olive": "Verde Oliva",
    "yellow": "Amarillo", "champagne": "Champán"
}

# --------------------------------------
# Helper Functions
# --------------------------------------
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
# CORE LOGIC: Robust Label Merging (Updated V3)
# --------------------------------------
def merge_shipping_and_manufacturing_labels(shipping_pdf_bytes, manufacturing_pdf_bytes, order_dataframe):
    """
    Robust Merging V3: 
    1. Indexes Manufacturing labels by Buyer Name.
    2. Scans Shipping labels for those Names (Text + OCR fallback).
    3. Appends correct Manufacturing pages immediately after Shipping pages.
    4. Returns list of unmatched buyers for alerting.
    """
    try:
        # --- STEP 1: Index the Manufacturing Labels ---
        mfg_reader = PdfReader(manufacturing_pdf_bytes)
        mfg_map = {} # Key: Buyer Name (Upper), Value: List of PdfReader Page Objects
        
        current_mfg_page_idx = 0
        
        for _, row in order_dataframe.iterrows():
            raw_name = str(row['Buyer Name']).strip().upper()
            # Clean name slightly (remove extra spaces)
            raw_name = " ".join(raw_name.split())
            
            if raw_name not in mfg_map:
                mfg_map[raw_name] = []
            
            # If we still have pages left, assign this one to the buyer
            if current_mfg_page_idx < len(mfg_reader.pages):
                mfg_map[raw_name].append(mfg_reader.pages[current_mfg_page_idx])
                current_mfg_page_idx += 1

        # Known buyers list for searching
        known_buyers = list(mfg_map.keys())

        # --- STEP 2: Process Shipping Labels ---
        output_pdf = PdfWriter()
        shipping_pdf_bytes.seek(0)
        
        processed_count = 0
        matched_count = 0
        
        with pdfplumber.open(shipping_pdf_bytes) as plist:
            ship_reader = PdfReader(shipping_pdf_bytes)
            
            for i, page in enumerate(plist.pages):
                # A. Extract Text
                text = page.extract_text() or ""
                text = text.upper()
                
                found_name = None
                
                # Strategy 1: Look for "SHIP TO" specific block (High Confidence)
                ship_to_match = re.search(r"SHIP\s*TO:?\s*\n+([^\n]+)", text)
                if ship_to_match:
                    candidate = ship_to_match.group(1).strip()
                    matches = get_close_matches(candidate, known_buyers, n=1, cutoff=0.8)
                    if matches:
                        found_name = matches[0]

                # Strategy 2: Scan full text for known buyers
                if not found_name:
                    for buyer in known_buyers:
                        if buyer in text:
                            found_name = buyer
                            break
                
                # Strategy 3: OCR Fallback (If page is an image/scan)
                if not found_name and len(text) < 50: 
                    try:
                        # Convert this specific page to image
                        images = convert_from_bytes(
                            shipping_pdf_bytes.getvalue(), 
                            first_page=i+1, 
                            last_page=i+1,
                            dpi=150
                        )
                        if images:
                            ocr_text = pytesseract.image_to_string(images[0]).upper()
                            for buyer in known_buyers:
                                if buyer in ocr_text:
                                    found_name = buyer
                                    break
                    except Exception as e:
                        print(f"OCR failed for page {i}: {e}")

                # --- STEP 3: Construct Output ---
                # Add the Shipping Label
                output_pdf.add_page(ship_reader.pages[i])
                processed_count += 1
                
                # Add the Manufacturing Label(s) if matched
                if found_name and found_name in mfg_map:
                    pages_to_add = mfg_map[found_name]
                    for p in pages_to_add:
                        output_pdf.add_page(p)
                        matched_count += 1
                    
                    # Remove from map so we know it was used
                    del mfg_map[found_name]

        # --- STEP 4: Handle Orphans & Reporting ---
        # Collect any buyers that remain in mfg_map (these are missing shipping labels)
        unmatched_buyers = list(mfg_map.keys())
        
        # Append orphans to the end so they aren't lost
        if len(mfg_map) > 0:
            for buyer, pages in mfg_map.items():
                for p in pages:
                    output_pdf.add_page(p)

        output_buffer = BytesIO()
        output_pdf.write(output_buffer)
        output_buffer.seek(0)
        
        return output_buffer, processed_count, matched_count, unmatched_buyers

    except Exception as e:
        st.error(f"Merge Error: {str(e)}")
        return BytesIO(), 0, 0, []

# --------------------------------------
# Airtable Functions
# --------------------------------------
def get_existing_order_ids():
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
            if offset: params["offset"] = offset
            response = requests.get(url, headers=headers, params=params)
            if response.status_code == 200:
                data = response.json()
                for record in data.get("records", []):
                    order_id = record.get("fields", {}).get("Order ID")
                    if order_id: existing_orders.add(order_id)
                offset = data.get("offset")
                if not offset: break
            else:
                break
    except Exception:
        pass
    return existing_orders

def upload_to_airtable(dataframe):
    headers = {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json"
    }
    
    st.info("🔍 Checking for duplicate orders...")
    existing_order_ids = get_existing_order_ids()
    
    unique_orders = dataframe[['Order ID', 'Order Date', 'Buyer Name']].drop_duplicates(subset=['Order ID'])
    new_orders = unique_orders[~unique_orders['Order ID'].isin(existing_order_ids)]
    
    if len(new_orders) == 0:
        st.info("ℹ️ All orders already exist in Airtable.")
        return 0, 0, []
    
    orders_created = 0
    line_items_created = 0
    errors = []
    progress_bar = st.progress(0)
    total_orders = len(new_orders)
    
    for idx, (_, order_row) in enumerate(new_orders.iterrows()):
        order_id = order_row['Order ID']
        try:
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
            response = requests.post(f"https://api.airtable.com/v0/{BASE_ID}/{ORDERS_TABLE}", headers=headers, json=order_payload)
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
                    item_response = requests.post(f"https://api.airtable.com/v0/{BASE_ID}/{LINE_ITEMS_TABLE}", headers=headers, json=line_item_payload)
                    if item_response.status_code == 200:
                        line_items_created += 1
            else:
                errors.append(f"Error creating order {order_id}")
        except Exception as e:
            errors.append(str(e))
        progress_bar.progress((idx + 1) / total_orders)
        
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
    st.markdown("### Version 10.3 (Alerts Added)")
    st.markdown("---")
    
    st.markdown("#### 📋 Quick Navigation")
    st.markdown('<a href="#upload-order" class="nav-link">📄 Upload Order</a>', unsafe_allow_html=True)
    st.markdown('<a href="#dashboard" class="nav-link">📊 Dashboard</a>', unsafe_allow_html=True)
    st.markdown('<a href="#color-analytics" class="nav-link">🎨 Color Analytics</a>', unsafe_allow_html=True)
    st.markdown('<a href="#bobbin-setup" class="nav-link">🧵 Bobbin Setup</a>', unsafe_allow_html=True)
    st.markdown('<a href="#generate-labels" class="nav-link">📥 Generate Labels</a>', unsafe_allow_html=True)
    st.markdown('<a href="#label-merge" class="nav-link">🔄 Label Merge</a>', unsafe_allow_html=True)
    st.markdown('<a href="#airtable-sync" class="nav-link">☁️ Airtable Sync</a>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown('<div class="status-indicator"><div class="status-dot"></div><span>System Ready</span></div>', unsafe_allow_html=True)

# --------------------------------------
# MAIN CONTENT
# --------------------------------------
st.title("🧵 Amazon Blanket Order Manager")
st.markdown("**Professional order processing & label generation system**")
st.markdown("---")

# File Upload Section
st.markdown('<a id="upload-order"></a>', unsafe_allow_html=True)
st.markdown("## 📄 Upload Order")
uploaded = st.file_uploader("Drop your Amazon packing slip PDF here", type=["pdf"])

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
        if m_id: order_id = m_id.group(1).strip()
        m_date = re.search(r"Order Date:\s*([A-Za-z]{3,},?\s*[A-Za-z]+\s*\d{1,2},?\s*\d{4})", page_text)
        if m_date: order_date = m_date.group(1).strip()

        blocks = re.split(r"(?=Customizations:)", page_text)
        for block in blocks:
            if "Customizations:" not in block:
                continue

            qty_match = re.search(r"Quantity\s*\n\s*(\d+)", block)
            quantity = qty_match.group(1) if qty_match else "1"

            blanket_color = ""
            thread_color = ""
            b_match = re.search(r"Color:\s*([^\n]+)", block)
            if b_match: blanket_color = clean_text(b_match.group(1))
            t_match = re.search(r"Thread Color:\s*([^\n]+)", block, re.IGNORECASE)
            if t_match: thread_color = translate_thread_color(clean_text(t_match.group(1)))

            name_match = re.search(r"Name:\s*([^\n]+)", block)
            customization_name = clean_text(name_match.group(1)) if name_match else ""

            beanie = "YES" if re.search(r"Personalized Baby Beanie:\s*Yes", block, re.IGNORECASE) else "NO"
            gift_box = "YES" if re.search(r"Gift Box\s*&\s*Gift Card:\s*Yes", block, re.IGNORECASE) else "NO"
            gift_note = "YES" if re.search(r"Gift Message:", block, re.IGNORECASE) else "NO"

            gift_msg_match = re.search(
                r"Gift Message:\s*([\s\S]*?)(?=\n(?:Grand total|Returning your item|Visit|Quantity|Order Totals|$))",
                block, re.IGNORECASE
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

    # Calculate Stats
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

    # Dashboard
    st.markdown("---")
    st.markdown('<a id="dashboard"></a>', unsafe_allow_html=True)
    st.markdown("## 📊 Order Dashboard")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1: st.metric("Total Blankets", total_blankets)
    with col2: st.metric("Total Orders", total_orders)
    with col3: st.metric("Beanies", total_beanies)
    with col4: st.metric("Gift Boxes", gift_boxes_needed)
    with col5: st.metric("Gift Messages", gift_messages_needed)
    with col6: st.metric("Unique Colors", len(blanket_color_counts))
    col7, col8 = st.columns(2)
    with col7: st.metric("Blanket Only", orders_blanket_only)
    with col8: st.metric("With Beanie", orders_with_beanie)

    # Color Analytics
    st.markdown("---")
    st.markdown('<a id="color-analytics"></a>', unsafe_allow_html=True)
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

    # Bobbin Setup
    st.markdown("---")
    st.markdown('<a id="bobbin-setup"></a>', unsafe_allow_html=True)
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

    # Generate Labels
    st.markdown("---")
    st.markdown('<a id="generate-labels"></a>', unsafe_allow_html=True)
    st.markdown("## 📥 Generate & Download")
    
    if 'manufacturing_labels_buffer' not in st.session_state:
        st.session_state.manufacturing_labels_buffer = None
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📦 Manufacturing Labels", use_container_width=True):
            with st.spinner("Generating manufacturing labels..."):
                pdf_data = generate_manufacturing_labels(df)
                st.session_state.manufacturing_labels_buffer = pdf_data
            st.success("✅ Labels generated!")
            st.download_button("⬇️ Download Manufacturing Labels", data=pdf_data, file_name="Manufacturing_Labels.pdf", mime="application/pdf", use_container_width=True)
    
    with col2:
        gift_count = len(df[df['Gift Message'] != ""])
        if st.button(f"💌 Gift Messages ({gift_count})", use_container_width=True):
            with st.spinner("Generating gift message labels..."):
                pdf_data = generate_gift_message_labels(df)
            st.success("✅ Labels generated!")
            st.download_button("⬇️ Download Gift Message Labels", data=pdf_data, file_name="Gift_Message_Labels.pdf", mime="application/pdf", use_container_width=True)
    
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
            st.download_button("⬇️ Download Summary PDF", data=pdf_data, file_name="Daily_Summary_Report.pdf", mime="application/pdf", use_container_width=True)

    # Label Merging
    st.markdown("---")
    st.markdown('<a id="label-merge"></a>', unsafe_allow_html=True)
    st.markdown("## 🔄 Merge Shipping & Manufacturing Labels")
    st.info("**Note:** You must generate Manufacturing Labels (button above) before using this tool.")
    
    shipping_labels_upload = st.file_uploader("📤 Upload Shipping Labels PDF", type=["pdf"], key="shipping_labels")
    
    if shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        if st.button("🔀 Merge Labels Now", type="primary", use_container_width=True):
            with st.spinner("Merging labels using Smart Match & OCR..."):
                shipping_labels_upload.seek(0)
                st.session_state.manufacturing_labels_buffer.seek(0)
                
                merged_pdf, num_shipping, num_manufacturing, unmatched_list = merge_shipping_and_manufacturing_labels(
                    shipping_labels_upload,
                    st.session_state.manufacturing_labels_buffer,
                    df
                )
                
                if merged_pdf:
                    st.success(f"✅ Processed {num_shipping} shipping labels and matched {num_manufacturing} manufacturing labels!")
                    
                    if unmatched_list:
                        st.error(f"⚠️ WARNING: {len(unmatched_list)} Orders did NOT find a matching Shipping Label!")
                        with st.expander("🚨 See Missing Orders (Attached at end of PDF)", expanded=True):
                            for name in unmatched_list:
                                st.write(f"❌ **{name}** - (No matching shipping label found)")
                    else:
                        st.balloons()

                    st.download_button("⬇️ Download Merged PDF", data=merged_pdf, file_name="Merged_Labels.pdf", mime="application/pdf", use_container_width=True)
    elif shipping_labels_upload and not st.session_state.manufacturing_labels_buffer:
        st.warning("⚠️ Please click 'Manufacturing Labels' button above first.")

    # Airtable Sync
    st.markdown("---")
    st.markdown('<a id="airtable-sync"></a>', unsafe_allow_html=True)
    st.markdown("## ☁️ Airtable Integration")
    if st.button("🚀 Upload to Airtable", type="primary", use_container_width=True):
        with st.spinner("Uploading to Airtable..."):
            orders_created, line_items_created, errors = upload_to_airtable(df)
        if errors:
            st.error(f"⚠️ Completed with errors: {len(errors)}")
            with st.expander("View Errors"):
                for error in errors: st.write(f"• {error}")
        else:
            st.success(f"✅ Successfully uploaded {orders_created} orders!")
