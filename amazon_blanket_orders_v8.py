"""
Amazon Blanket Orders Parser v8.0
- Parses baby blanket orders from Amazon PDFs
- Generates 4×6 production labels + gift note labels
- Airtable integration with duplicate prevention
- Auto-skips towel orders
"""

import io
import re
from dataclasses import dataclass, field
from typing import List, Dict, Optional
from datetime import datetime

import streamlit as st
import pdfplumber
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape

# ======================================================
# AIRTABLE CONFIGURATION
# ======================================================
# Install: pip install pyairtable
try:
    from pyairtable import Api
    AIRTABLE_AVAILABLE = True
except ImportError:
    AIRTABLE_AVAILABLE = False

# 🔧 CONFIGURATION - Add your Airtable credentials here
AIRTABLE_API_KEY = ""  # Get from: https://airtable.com/account
AIRTABLE_BASE_ID = ""  # Found in your Airtable URL
AIRTABLE_TABLE_NAME = "amazon blanket orders"

# ======================================================
# CONSTANTS
# ======================================================
st.set_page_config(page_title="Amazon Blanket Orders v8.0", layout="wide")

# Thread color Spanish translations
THREAD_COLOR_ES = {
    "White": "Blanco", "Black": "Negro", "Gold": "Dorado", "Silver": "Plateado",
    "Red": "Rojo", "Blue": "Azul", "Navy": "Azul Marino", "Navy Blue": "Azul Marino",
    "Light Blue": "Azul Claro", "Green": "Verde", "Pink": "Rosa", "Hot Pink": "Rosa Fucsia",
    "Lilac": "Lila", "Purple": "Morado", "Yellow": "Amarillo", "Beige": "Beige",
    "Brown": "Marrón", "Gray": "Gris", "Grey": "Gris", "Orange": "Naranja",
    "Teal": "Verde Azulado", "Ivory": "Marfil", "Champagne": "Champán",
    "Dark Grey": "Gris Oscuro", "Salmon Pink": "Rosa Salmón", "Light Pink": "Rosa Claro",
}

# Towel product types (will be detected but skipped)
TOWEL_SKUS = ["Set-3Pcs", "Set-6Pcs", "HT-2PCS", "BT-2Pcs", "BS-1Pcs"]

# Regex patterns
ORDER_ID_REGEX = re.compile(r"\bOrder ID[:\s]+([0-9\-]+)", re.IGNORECASE)
ORDER_DATE_REGEX = re.compile(r"\bOrder Date[:\s]+([A-Za-z]{3,},?\s*[A-Za-z]+\s*\d{1,2},?\s*\d{4})", re.IGNORECASE)
SHIPPING_SERVICE_REGEX = re.compile(r"\bShipping Service[:\s]+([A-Za-z\s]+)", re.IGNORECASE)
BUYER_NAME_REGEX = re.compile(r"\bShip To:\s*(.+)", re.IGNORECASE)

# ======================================================
# DATA CLASS
# ======================================================
@dataclass
class BlanketOrder:
    order_id: str = ""
    order_date: str = ""
    buyer_name: str = ""
    shipping_service: str = "Standard"
    blanket_color: str = ""
    thread_color: str = ""
    thread_color_es: str = ""
    embroidery_font: str = ""
    embroidery_length: str = ""
    name: str = ""
    quantity: int = 1
    beanie: str = "NO"
    gift_box: str = "NO"
    gift_note: str = "NO"
    gift_message: str = ""
    is_towel: bool = False
    
    def to_dict(self):
        """Convert to dictionary for table display"""
        return {
            "Order ID": self.order_id,
            "Order Date": self.order_date,
            "Buyer Name": self.buyer_name,
            "Quantity": self.quantity,
            "Blanket Color": self.blanket_color,
            "Thread Color": f"{self.thread_color.upper()} ({self.thread_color_es.upper()})" if self.thread_color_es else self.thread_color.upper(),
            "Name": self.name,
            "Font": self.embroidery_font,
            "Length": self.embroidery_length,
            "Beanie": self.beanie,
            "Gift Box": self.gift_box,
            "Gift Note": self.gift_note,
            "Gift Message": self.gift_message[:50] + "..." if len(self.gift_message) > 50 else self.gift_message,
            "Shipping": self.shipping_service,
        }

# ======================================================
# HELPER FUNCTIONS
# ======================================================
def clean_text(text: str) -> str:
    """Clean extracted text"""
    if not text:
        return ""
    text = re.sub(r"\(#?[A-Fa-f0-9]{3,6}\)", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def get_thread_color_spanish(color: str) -> str:
    """Get Spanish translation for thread color"""
    if not color:
        return ""
    color_clean = color.strip().title()
    for eng, esp in THREAD_COLOR_ES.items():
        if eng.lower() in color_clean.lower():
            return esp
    return ""

def is_towel_order(text: str) -> bool:
    """Check if order contains towel SKUs"""
    for sku in TOWEL_SKUS:
        if sku in text:
            return True
    return False

# ======================================================
# PDF PARSER
# ======================================================
def parse_blanket_pdf(uploaded_file) -> tuple[List[BlanketOrder], int]:
    """
    Parse Amazon blanket order PDF
    Returns: (list of orders, count of skipped towel orders)
    """
    orders = []
    towel_count = 0
    
    with pdfplumber.open(uploaded_file) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"
        
        # Better splitting: Use "Shipping Address:" as primary delimiter
        # This ensures each order block is complete
        order_blocks = re.split(r"(?=Shipping Address:)", full_text)
        
        for block in order_blocks:
            if "Order ID" not in block or "Shipping Address:" not in block:
                continue
            
            # Check if this is a towel order
            if is_towel_order(block):
                towel_count += 1
                continue
            
            # Check if this is a blanket order
            if "patique Personalized Baby Blanket" not in block and "VC-R4JI-YQED" not in block:
                continue
            
            # Extract order metadata ONCE for the entire order
            order_id = ""
            order_date = ""
            buyer_name = ""
            shipping_service = "Standard"
            
            order_id_match = ORDER_ID_REGEX.search(block)
            if order_id_match:
                order_id = order_id_match.group(1).strip()
            
            date_match = ORDER_DATE_REGEX.search(block)
            if date_match:
                order_date = date_match.group(1).strip()
            
            # Extract buyer name from "Ship To:" section
            buyer_match = BUYER_NAME_REGEX.search(block)
            if buyer_match:
                # Get the text after "Ship To:" and before the address lines
                ship_to_text = buyer_match.group(1)
                lines = [l.strip() for l in ship_to_text.split("\n") if l.strip()]
                if lines:
                    # First non-empty line is the buyer name
                    buyer_name = lines[0]
                    # Stop at address numbers/street
                    if re.match(r'^\d+\s+', buyer_name):
                        # If first line starts with number, it's an address - look for name before
                        name_match = re.search(r"Ship To:\s*([^\n]+)", block)
                        if name_match:
                            buyer_name = name_match.group(1).strip()
            
            shipping_match = SHIPPING_SERVICE_REGEX.search(block)
            if shipping_match:
                shipping_service = shipping_match.group(1).strip()
            
            # Find all item sections within this order
            # Split by "Order Item ID:" to get individual items
            item_sections = re.split(r"(?=Order Item ID:)", block)
            
            for section in item_sections:
                # Must have both Order Item ID and Customizations
                if "Order Item ID:" not in section or "Customizations:" not in section:
                    continue
                
                # Make sure this section is still part of blanket order
                if "patique Personalized Baby Blanket" not in section and "VC-R4JI-YQED" not in section:
                    continue
                
                order = BlanketOrder()
                
                # Apply order metadata to each item
                order.order_id = order_id
                order.order_date = order_date
                order.buyer_name = buyer_name
                order.shipping_service = shipping_service
                
                # Extract quantity from this specific item
                qty_match = re.search(r"Quantity.*?(\d+)", section, re.IGNORECASE | re.DOTALL)
                if qty_match:
                    order.quantity = int(qty_match.group(1))
                
                # Get customization section
                custom_split = section.split("Customizations:")
                if len(custom_split) < 2:
                    continue
                    
                custom_section = custom_split[1]
                
                # Blanket color
                color_match = re.search(r"Color:\s*([^\n]+)", custom_section)
                if color_match:
                    order.blanket_color = clean_text(color_match.group(1))
                
                # Thread color
                thread_match = re.search(r"Thread Color:\s*([^\n]+)", custom_section, re.IGNORECASE)
                if thread_match:
                    thread_raw = clean_text(thread_match.group(1))
                    order.thread_color = thread_raw
                    order.thread_color_es = get_thread_color_spanish(thread_raw)
                
                # Embroidery font
                font_match = re.search(r"Embroidery Font:\s*([^\n]+)", custom_section, re.IGNORECASE)
                if font_match:
                    order.embroidery_font = clean_text(font_match.group(1))
                
                # Embroidery length
                length_match = re.search(r"Choose Embroidery Length:\s*([^\n]+)", custom_section, re.IGNORECASE)
                if length_match:
                    order.embroidery_length = clean_text(length_match.group(1))
                
                # Name
                name_match = re.search(r"Name:\s*([^\n]+)", custom_section)
                if name_match:
                    order.name = clean_text(name_match.group(1)).upper()
                
                # Beanie
                if re.search(r"Personalized Baby Beanie:\s*Yes", custom_section, re.IGNORECASE):
                    order.beanie = "YES"
                else:
                    order.beanie = "NO"
                
                # Gift box
                if re.search(r"Gift Box.*?Yes", custom_section, re.IGNORECASE):
                    order.gift_box = "YES"
                else:
                    order.gift_box = "NO"
                
                # Gift message - only within this item's section
                gift_match = re.search(r"Gift Message:\s*(.*?)(?=\n(?:Item subtotal|Grand total|Returning|Quantity|Order Item ID|$))", custom_section, re.IGNORECASE | re.DOTALL)
                if gift_match:
                    order.gift_message = clean_text(gift_match.group(1))
                    order.gift_note = "YES"
                else:
                    order.gift_note = "NO"
                
                # Only add if we have minimum required fields
                if order.order_id and order.buyer_name:
                    orders.append(order)
    
    return orders, towel_count

# ======================================================
# LABEL GENERATION - PRODUCTION LABELS
# ======================================================
def generate_production_labels(orders: List[BlanketOrder]) -> bytes:
    """Generate 4×6 landscape production labels"""
    buf = io.BytesIO()
    PAGE_W, PAGE_H = 6 * inch, 4 * inch
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    
    for order in orders:
        x0 = 0.3 * inch
        y = PAGE_H - 0.35 * inch
        
        # Header - Order ID and Buyer
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x0, y, f"ORDER: {order.order_id}")
        y -= 0.22 * inch
        
        c.setFont("Helvetica-Bold", 15)
        c.drawString(x0, y, order.buyer_name[:45])
        y -= 0.28 * inch
        
        # Main info boxes - Blanket Color and Quantity
        box_width = (PAGE_W - 0.6 * inch - 0.15 * inch) / 2
        
        # Blanket Color box
        c.setFillColorRGB(0.93, 0.94, 0.97)
        c.rect(x0, y - 0.4 * inch, box_width, 0.42 * inch, fill=1, stroke=0)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", 10)
        c.drawString(x0 + 0.08 * inch, y - 0.12 * inch, "Blanket Color")
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x0 + 0.08 * inch, y - 0.3 * inch, order.blanket_color[:18])
        
        # Quantity box
        c.setFillColorRGB(0.93, 0.94, 0.97)
        c.rect(x0 + box_width + 0.15 * inch, y - 0.4 * inch, box_width, 0.42 * inch, fill=1, stroke=0)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", 10)
        c.drawString(x0 + box_width + 0.23 * inch, y - 0.12 * inch, "Quantity")
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x0 + box_width + 0.23 * inch, y - 0.3 * inch, str(order.quantity))
        
        y -= 0.52 * inch
        
        # Thread Color - Full width, highlighted
        c.setFillColorRGB(1, 0.96, 0.96)
        c.rect(x0, y - 0.38 * inch, PAGE_W - 0.6 * inch, 0.42 * inch, fill=1, stroke=0)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", 10)
        c.drawString(x0 + 0.08 * inch, y - 0.12 * inch, "Thread Color")
        c.setFont("Helvetica-Bold", 13)
        thread_display = f"{order.thread_color.upper()} ({order.thread_color_es.upper()})" if order.thread_color_es else order.thread_color.upper()
        # Truncate if too long
        if len(thread_display) > 40:
            thread_display = thread_display[:40] + "..."
        c.setFillColorRGB(0.77, 0.19, 0.19)
        c.drawString(x0 + 0.08 * inch, y - 0.3 * inch, thread_display)
        c.setFillColorRGB(0, 0, 0)
        
        y -= 0.48 * inch
        
        # Extras row - Beanie, Gift Box, Gift Note
        extra_width = (PAGE_W - 0.6 * inch - 0.3 * inch) / 3
        
        for i, (label, value) in enumerate([
            ("Beanie", order.beanie),
            ("Gift Box", order.gift_box),
            ("Gift Note", order.gift_note)
        ]):
            box_x = x0 + i * (extra_width + 0.15 * inch)
            c.setFillColorRGB(0.9, 1, 0.98)
            c.rect(box_x, y - 0.35 * inch, extra_width, 0.38 * inch, fill=1, stroke=0)
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 9)
            c.drawString(box_x + 0.05 * inch, y - 0.1 * inch, label)
            c.setFont("Helvetica-Bold", 12)
            c.setFillColorRGB(0.14, 0.31, 0.32)
            c.drawString(box_x + 0.05 * inch, y - 0.27 * inch, value)
            c.setFillColorRGB(0, 0, 0)
        
        y -= 0.47 * inch
        
        # Personalization section
        c.setFillColorRGB(1, 0.98, 0.92)
        c.rect(x0, y - 0.7 * inch, PAGE_W - 0.6 * inch, 0.75 * inch, fill=1, stroke=0)
        c.setFillColorRGB(0, 0, 0)
        
        c.setFont("Helvetica-Bold", 12)
        c.setFillColorRGB(0.46, 0.26, 0.06)
        c.drawString(x0 + 0.08 * inch, y - 0.14 * inch, "⭐ PERSONALIZATION ⭐")
        c.setFillColorRGB(0, 0, 0)
        
        c.setFont("Helvetica", 13)
        c.drawString(x0 + 0.08 * inch, y - 0.34 * inch, f"Name: ")
        c.setFont("Helvetica-Bold", 13)
        # Truncate name if too long
        name_display = order.name if len(order.name) <= 35 else order.name[:35] + "..."
        c.drawString(x0 + 0.55 * inch, y - 0.34 * inch, name_display)
        
        c.setFont("Helvetica", 13)
        c.drawString(x0 + 0.08 * inch, y - 0.52 * inch, f"Font: ")
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x0 + 0.45 * inch, y - 0.52 * inch, order.embroidery_font[:20])
        
        c.setFont("Helvetica", 13)
        c.drawString(x0 + 0.08 * inch, y - 0.68 * inch, f"Length: ")
        c.setFont("Helvetica-Bold", 13)
        # Truncate length if needed
        length_display = order.embroidery_length if len(order.embroidery_length) <= 30 else order.embroidery_length[:30]
        c.drawString(x0 + 0.55 * inch, y - 0.68 * inch, length_display)
        
        c.showPage()
    
    c.save()
    return buf.getvalue()

# ======================================================
# LABEL GENERATION - GIFT NOTE LABELS
# ======================================================
def generate_gift_note_labels(orders: List[BlanketOrder]) -> bytes:
    """Generate 4×6 landscape gift note labels (only for orders with messages)"""
    buf = io.BytesIO()
    PAGE_W, PAGE_H = 6 * inch, 4 * inch
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    
    gift_orders = [o for o in orders if o.gift_message]
    
    if not gift_orders:
        # Create a blank page if no gift messages
        c.setFont("Helvetica", 12)
        c.drawString(2 * inch, 2 * inch, "No gift messages to print")
        c.showPage()
    else:
        for order in gift_orders:
            # Draw double border frame
            c.setStrokeColor(colors.HexColor("#8B4513"))
            c.setLineWidth(8)
            c.rect(0.08 * inch, 0.08 * inch, PAGE_W - 0.16 * inch, PAGE_H - 0.16 * inch, stroke=1, fill=0)
            
            # Inner border
            c.setStrokeColor(colors.HexColor("#D4AF37"))
            c.setLineWidth(2)
            c.roundRect(0.23 * inch, 0.23 * inch, PAGE_W - 0.46 * inch, PAGE_H - 0.46 * inch, 8, stroke=1, fill=0)
            
            # Corner decorations
            corner_size = 0.3 * inch
            corners = [
                (0.16 * inch, PAGE_H - 0.16 * inch - corner_size, 0.16 * inch + corner_size, PAGE_H - 0.16 * inch),  # top-left
                (PAGE_W - 0.16 * inch - corner_size, PAGE_H - 0.16 * inch - corner_size, PAGE_W - 0.16 * inch, PAGE_H - 0.16 * inch),  # top-right
                (0.16 * inch, 0.16 * inch, 0.16 * inch + corner_size, 0.16 * inch + corner_size),  # bottom-left
                (PAGE_W - 0.16 * inch - corner_size, 0.16 * inch, PAGE_W - 0.16 * inch, 0.16 * inch + corner_size),  # bottom-right
            ]
            
            c.setStrokeColor(colors.HexColor("#D4AF37"))
            c.setLineWidth(3)
            for x1, y1, x2, y2 in corners:
                if x1 < PAGE_W / 2 and y1 > PAGE_H / 2:  # top-left
                    c.line(x1, y2, x2, y2)
                    c.line(x1, y1, x1, y2)
                elif x1 > PAGE_W / 2 and y1 > PAGE_H / 2:  # top-right
                    c.line(x1, y2, x2, y2)
                    c.line(x2, y1, x2, y2)
                elif x1 < PAGE_W / 2 and y1 < PAGE_H / 2:  # bottom-left
                    c.line(x1, y1, x2, y1)
                    c.line(x1, y1, x1, y2)
                else:  # bottom-right
                    c.line(x1, y1, x2, y1)
                    c.line(x2, y1, x2, y2)
            
            # Gift message - centered
            c.setFont("Times-Italic", 18)
            c.setFillColor(colors.HexColor("#2d3748"))
            
            # Word wrap the message
            words = order.gift_message.split()
            lines = []
            current_line = []
            max_width = PAGE_W - 1.5 * inch
            
            for word in words:
                test_line = ' '.join(current_line + [word])
                if c.stringWidth(test_line, "Times-Italic", 18) < max_width:
                    current_line.append(word)
                else:
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [word]
            if current_line:
                lines.append(' '.join(current_line))
            
            # Center vertically
            total_height = len(lines) * 0.3 * inch
            y_start = (PAGE_H + total_height) / 2
            
            for line in lines:
                text_width = c.stringWidth(line, "Times-Italic", 18)
                x = (PAGE_W - text_width) / 2
                c.drawString(x, y_start, line)
                y_start -= 0.3 * inch
            
            c.showPage()
    
    c.save()
    return buf.getvalue()

# ======================================================
# AIRTABLE FUNCTIONS
# ======================================================
def get_existing_order_ids(api_key: str, base_id: str, table_name: str) -> set:
    """Get all existing Order IDs from Airtable"""
    if not AIRTABLE_AVAILABLE:
        return set()
    
    try:
        api = Api(api_key)
        table = api.table(base_id, table_name)
        records = table.all()
        return {record['fields'].get('Order ID', '') for record in records if record['fields'].get('Order ID')}
    except Exception as e:
        st.error(f"Error reading from Airtable: {str(e)}")
        return set()

def send_to_airtable(orders: List[BlanketOrder], api_key: str, base_id: str, table_name: str) -> tuple[int, int, List[str]]:
    """
    Send orders to Airtable with duplicate checking
    Returns: (added_count, skipped_count, skipped_order_ids)
    """
    if not AIRTABLE_AVAILABLE:
        raise Exception("pyairtable not installed. Run: pip install pyairtable")
    
    # Get existing order IDs
    existing_ids = get_existing_order_ids(api_key, base_id, table_name)
    
    # Separate new and duplicate orders
    new_orders = [o for o in orders if o.order_id not in existing_ids]
    duplicate_orders = [o for o in orders if o.order_id in existing_ids]
    
    # Add new orders to Airtable
    added_count = 0
    if new_orders:
        try:
            api = Api(api_key)
            table = api.table(base_id, table_name)
            
            for order in new_orders:
                record = {
                    "Order ID": order.order_id,
                    "Order Date": order.order_date,
                    "Buyer Name": order.buyer_name,
                    "Shipping Service": order.shipping_service,
                    "Blanket Color": order.blanket_color,
                    "Thread Color": order.thread_color,
                    "Thread Color (Spanish)": order.thread_color_es,
                    "Embroidery Font": order.embroidery_font,
                    "Embroidery Length": order.embroidery_length,
                    "Name": order.name,
                    "Quantity": order.quantity,
                    "Beanie": order.beanie,
                    "Gift Box": order.gift_box,
                    "Gift Note": order.gift_note,
                    "Gift Message": order.gift_message,
                    "Status": "New",
                    "Date Added": datetime.now().isoformat(),
                }
                table.create(record)
                added_count += 1
        except Exception as e:
            st.error(f"Error writing to Airtable: {str(e)}")
    
    skipped_ids = [o.order_id for o in duplicate_orders]
    return added_count, len(duplicate_orders), skipped_ids

# ======================================================
# STREAMLIT UI
# ======================================================
def main():
    st.title("🧵 Amazon Blanket Orders v8.0")
    st.caption("Parse PDFs → Generate Labels → Send to Airtable")
    
    # Sidebar - Airtable Configuration
    with st.sidebar:
        st.header("⚙️ Airtable Configuration")
        
        if not AIRTABLE_AVAILABLE:
            st.error("⚠️ Airtable not installed")
            st.code("pip install pyairtable")
        else:
            st.success("✅ Airtable library loaded")
        
        api_key = st.text_input("API Key", value=AIRTABLE_API_KEY, type="password")
        base_id = st.text_input("Base ID", value=AIRTABLE_BASE_ID)
        table_name = st.text_input("Table Name", value=AIRTABLE_TABLE_NAME)
        
        airtable_configured = bool(api_key and base_id and table_name)
        
        if airtable_configured:
            st.success("✅ Airtable configured")
        else:
            st.warning("⚠️ Add Airtable credentials to enable sync")
        
        st.divider()
        st.markdown("""
        ### 📚 Instructions
        1. Upload Amazon PDF
        2. Review parsed data
        3. Generate labels
        4. (Optional) Send to Airtable
        """)
    
    # Main content
    st.header("📄 Step 1: Upload PDF")
    uploaded_file = st.file_uploader("Upload Amazon order PDF", type=["pdf"])
    
    if uploaded_file:
        with st.spinner("🔍 Parsing PDF..."):
            orders, towel_count = parse_blanket_pdf(uploaded_file)
        
        if towel_count > 0:
            st.warning(f"⚠️ Detected and skipped {towel_count} towel order(s)")
        
        if not orders:
            st.error("❌ No blanket orders found in PDF")
            return
        
        st.success(f"✅ Found {len(orders)} blanket order(s)")
        
        # Display parsed data
        st.header("📊 Step 2: Review Parsed Data")
        df = pd.DataFrame([o.to_dict() for o in orders])
        st.dataframe(df, use_container_width=True)
        
        # CSV Export
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f"blanket_orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
        
        # Label Generation
        st.header("🏷️ Step 3: Generate Labels")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📦 Generate Production Labels", type="primary", use_container_width=True):
                with st.spinner("Generating production labels..."):
                    pdf_bytes = generate_production_labels(orders)
                st.download_button(
                    label="⬇️ Download Production Labels (4×6)",
                    data=pdf_bytes,
                    file_name=f"production_labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        
        with col2:
            gift_orders = [o for o in orders if o.gift_message]
            if gift_orders:
                if st.button(f"💝 Generate Gift Note Labels ({len(gift_orders)})", type="primary", use_container_width=True):
                    with st.spinner("Generating gift note labels..."):
                        pdf_bytes = generate_gift_note_labels(orders)
                    st.download_button(
                        label="⬇️ Download Gift Note Labels (4×6)",
                        data=pdf_bytes,
                        file_name=f"gift_notes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            else:
                st.info("ℹ️ No gift messages in this batch")
        
        # Airtable Integration
        if airtable_configured:
            st.header("📤 Step 4: Send to Airtable (Optional)")
            
            if 'show_airtable_preview' not in st.session_state:
                st.session_state.show_airtable_preview = False
            
            if not st.session_state.show_airtable_preview:
                if st.button("📤 Send to Airtable", type="secondary", use_container_width=True):
                    st.session_state.show_airtable_preview = True
                    st.rerun()
            
            if st.session_state.show_airtable_preview:
                # Check for duplicates
                with st.spinner("Checking for duplicates..."):
                    existing_ids = get_existing_order_ids(api_key, base_id, table_name)
                
                new_orders = [o for o in orders if o.order_id not in existing_ids]
                duplicate_orders = [o for o in orders if o.order_id in existing_ids]
                
                # Show preview
                st.subheader("📋 Airtable Upload Preview")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("✅ NEW Orders (will be added)", len(new_orders))
                with col2:
                    st.metric("⚠️ DUPLICATES (will be skipped)", len(duplicate_orders))
                
                if duplicate_orders:
                    with st.expander("📝 View Duplicate Order IDs"):
                        for order in duplicate_orders:
                            st.text(f"• {order.order_id} - {order.buyer_name}")
                
                st.divider()
                
                # Confirmation buttons
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("✅ Confirm Upload", type="primary", use_container_width=True):
                        if new_orders:
                            with st.spinner("Uploading to Airtable..."):
                                added, skipped, skipped_ids = send_to_airtable(orders, api_key, base_id, table_name)
                            st.success(f"✅ Successfully added {added} order(s) to Airtable!")
                            if skipped > 0:
                                st.info(f"ℹ️ Skipped {skipped} duplicate(s)")
                        else:
                            st.warning("⚠️ All orders already exist in Airtable - nothing to add")
                        st.session_state.show_airtable_preview = False
                
                with col2:
                    if st.button("❌ Cancel", use_container_width=True):
                        st.session_state.show_airtable_preview = False
                        st.rerun()
        
        else:
            st.header("📤 Step 4: Airtable Integration")
            st.info("👈 Configure Airtable credentials in the sidebar to enable upload")

if __name__ == "__main__":
    main()
