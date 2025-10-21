import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch, landscape, portrait
from reportlab.lib import colors
import requests

# --------------------------------------
# Airtable Configuration
# --------------------------------------
AIRTABLE_PAT = "pat3HPlu7bZzJep6t.2ea662c7b5e4f25f406969f987c5fdb9e15d5a2c6e933428934c8e5602ae7a68"
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

# --------------------------------------
# Airtable Functions
# --------------------------------------
def upload_to_airtable(dataframe):
    """Upload parsed orders to Airtable"""
    headers = {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json"
    }
    
    unique_orders = dataframe[['Order ID', 'Order Date', 'Buyer Name']].drop_duplicates(subset=['Order ID'])
    
    orders_created = 0
    line_items_created = 0
    errors = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_orders = len(unique_orders)
    
    for idx, (_, order_row) in enumerate(unique_orders.iterrows()):
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
                "Buyer Name": order_row['Buyer Name'],  # FIRST FIELD
                "Order ID": [airtable_order_id],
                "Customization Name Placement": item_row['Customization Name'],
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
# Streamlit Setup
# --------------------------------------
st.set_page_config(page_title="Amazon Blanket Orders – v9.0", layout="centered")
st.title("🧵 Amazon Blanket Order Manager — v9.0")

st.write("""
### 🪡 Features
- **Parse Amazon PDFs** and generate manufacturing labels
- **Upload to Airtable** for order tracking & team management
- **Two-column layout** with smart text wrapping
- **End-of-day summary** with accurate order counts
- **Bobbin color grouping** for embroidery setup
""")

uploaded = st.file_uploader("📄 Upload your Amazon packing slip PDF", type=["pdf"])

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
    st.success(f"✅ {len(df)} line items parsed from {df['Order ID'].nunique()} orders")
    st.dataframe(df)

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
    
    black_bobbin_threads = df[df['Bobbin_Color'] == 'Black Bobbin'].groupby('Thread Color')['Quantity_Int'].sum().sort_values(ascending=False)
    white_bobbin_threads = df[df['Bobbin_Color'] == 'White Bobbin'].groupby('Thread Color')['Quantity_Int'].sum().sort_values(ascending=False)

    # --------------------------------------
    # Display Summary
    # --------------------------------------
    st.write("---")
    st.header("📊 End of Day Summary")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("🧵 Total Blankets", total_blankets)
    with col2:
        st.metric("🧢 Total Beanies", total_beanies)
    with col3:
        st.metric("📦 Total Orders", total_orders)
    with col4:
        st.metric("💌 Gift Messages", gift_messages_needed)
    
    col5, col6, col7, col8 = st.columns(4)
    
    with col5:
        st.metric("🎁 Gift Boxes", gift_boxes_needed)
    with col6:
        st.metric("Blanket Only", orders_blanket_only)
    with col7:
        st.metric("With Beanie", orders_with_beanie)
    with col8:
        st.metric("Unique Blanket Colors", len(blanket_color_counts))
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("🎨 Blanket Color Breakdown")
        for color, count in blanket_color_counts.items():
            st.write(f"**{color}:** {count}")
    
    with col_b:
        st.subheader("🧵 Thread Color Breakdown")
        for color, count in thread_color_counts.items():
            st.write(f"**{color}:** {count}")
    
    st.write("---")
    st.subheader("🎯 Bobbin Color Setup")
    
    col_bobbin1, col_bobbin2 = st.columns(2)
    
    with col_bobbin1:
        st.markdown("### ⚫ Black Bobbin")
        st.metric("Total Items", bobbin_counts.get('Black Bobbin', 0))
        if len(black_bobbin_threads) > 0:
            for color, count in black_bobbin_threads.items():
                st.write(f"• {color}: {count}")
        else:
            st.write("_No items_")
    
    with col_bobbin2:
        st.markdown("### ⚪ White Bobbin")
        st.metric("Total Items", bobbin_counts.get('White Bobbin', 0))
        if len(white_bobbin_threads) > 0:
            for color, count in white_bobbin_threads.items():
                st.write(f"• {color}: {count}")
        else:
            st.write("_No items_")

    # --------------------------------------
    # Airtable Upload Section
    # --------------------------------------
    st.write("---")
    st.header("☁️ Upload to Airtable")
    
    st.info("📤 Click below to upload these orders to your Airtable base for tracking and management.")
    
    if st.button("🚀 Upload to Airtable", type="primary", use_container_width=True):
        with st.spinner("Uploading to Airtable..."):
            orders_created, line_items_created, errors = upload_to_airtable(df)
        
        if errors:
            st.error(f"⚠️ Upload completed with {len(errors)} errors")
            with st.expander("View Errors"):
                for error in errors:
                    st.write(f"• {error}")
        else:
            st.success(f"✅ Successfully uploaded!")
        
        col_result1, col_result2 = st.columns(2)
        with col_result1:
            st.metric("Orders Created", orders_created)
        with col_result2:
            st.metric("Line Items Created", line_items_created)
        
        st.info("🔗 Go to your Airtable base to view and manage orders!")

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
        middle_x = left + (right - left) * 0.60

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

            box_height = 0.65 * inch
            box_y = y - box_height
            c.setStrokeColor(colors.black)
            c.setLineWidth(2)
            c.rect(left, box_y, right - left, box_height, stroke=1, fill=0)
            
            c.setFont("Helvetica-Bold", 16)
            text_y = box_y + box_height - 0.22 * inch
            c.drawString(left + 0.1 * inch, text_y, f"THREAD COLOR: {row['Thread Color']}")
            
            text_y -= 0.28 * inch
            c.setFont("Helvetica-Bold", 15)
            c.drawString(left + 0.1 * inch, text_y, f"BLANKET COLOR: {row['Blanket Color']}")
            y = box_y - 0.3 * inch

            c.setStrokeColor(colors.grey)
            c.setLineWidth(1)
            c.line(middle_x, y, middle_x, 0.3 * inch)

            y_left = y
            c.setFont("Helvetica-Bold", 15)
            c.drawString(left, y_left, "PRODUCT:")
            y_left -= 0.3 * inch

            c.setFont("Helvetica-Bold", 15)
            name_text = f"★ Name: {row['Customization Name']}"
            left_column_width = middle_x - left - 0.2 * inch
            
            if c.stringWidth(name_text, "Helvetica-Bold", 15) <= left_column_width:
                c.drawString(left, y_left, name_text)
                y_left -= 0.3 * inch
            else:
                c.drawString(left, y_left, "★ Name:")
                y_left -= 0.25 * inch
                name_only = row['Customization Name']
                
                words = name_only.split()
                if len(words) > 1 and c.stringWidth(name_only, "Helvetica-Bold", 15) > left_column_width - 0.3 * inch:
                    line1_words = []
                    line2_words = []
                    for word in words:
                        test_line1 = ' '.join(line1_words + [word])
                        if c.stringWidth(test_line1, "Helvetica-Bold", 15) <= left_column_width - 0.3 * inch:
                            line1_words.append(word)
                        else:
                            line2_words.append(word)
                    
                    c.drawString(left + 0.3 * inch, y_left, ' '.join(line1_words))
                    if line2_words:
                        y_left -= 0.22 * inch
                        c.drawString(left + 0.3 * inch, y_left, ' '.join(line2_words))
                else:
                    c.drawString(left + 0.3 * inch, y_left, name_only)
                
                y_left -= 0.3 * inch

            c.setFont("Helvetica-Bold", 14)
            checkbox = "☑" if row['Include Beanie'] == "YES" else "☐"
            c.drawString(left, y_left, f"{checkbox} Beanie: {row['Include Beanie']}")

            y_right = y
            c.setFont("Helvetica-Bold", 15)
            c.drawString(middle_x + 0.2 * inch, y_right, "PACKAGING:")
            y_right -= 0.3 * inch

            c.setFont("Helvetica-Bold", 14)
            checkbox = "☑" if row['Gift Box'] == "YES" else "☐"
            c.drawString(middle_x + 0.2 * inch, y_right, f"{checkbox} Gift Box: {row['Gift Box']}")
            y_right -= 0.3 * inch

            checkbox = "☑" if row['Gift Note'] == "YES" else "☐"
            c.drawString(middle_x + 0.2 * inch, y_right, f"{checkbox} Gift Note: {row['Gift Note']}")

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
                c.setFillColor(colors.white)
                c.rect(0, 0, W, H, fill=1, stroke=0)
                
                c.setStrokeColor(colors.HexColor("#FFB6C1"))
                c.setLineWidth(3)
                c.setDash([])
                c.rect(0.4 * inch, 0.4 * inch, W - 0.8 * inch, H - 0.8 * inch, stroke=1, fill=0)

                c.setFillColor(colors.HexColor("#4A4A4A"))
                c.setFont("Helvetica", 16)
                message = row['Gift Message']
                
                words = message.split()
                lines = []
                current_line = []
                max_width = W - 1.2 * inch
                
                for word in words:
                    test_line = ' '.join(current_line + [word])
                    if c.stringWidth(test_line, "Helvetica", 16) < max_width:
                        current_line.append(word)
                    else:
                        if current_line:
                            lines.append(' '.join(current_line))
                        current_line = [word]
                
                if current_line:
                    lines.append(' '.join(current_line))

                total_height = len(lines) * 0.28 * inch
                y = (H + total_height) / 2

                for line in lines:
                    c.drawCentredString(W / 2, y, line)
                    y -= 0.28 * inch

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
    # Download Buttons
    # --------------------------------------
    st.write("---")
    st.subheader("📥 Generate Labels & Reports")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📦 Manufacturing Labels", use_container_width=True):
            pdf_data = generate_manufacturing_labels(df)
            st.download_button(
                label="⬇️ Download Manufacturing Labels",
                data=pdf_data,
                file_name="Manufacturing_Labels.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col2:
        gift_count = len(df[df['Gift Message'] != ""])
        if st.button(f"💌 Gift Message Labels ({gift_count})", use_container_width=True):
            pdf_data = generate_gift_message_labels(df)
            st.download_button(
                label="⬇️ Download Gift Message Labels",
                data=pdf_data,
                file_name="Gift_Message_Labels.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with col3:
        if st.button("📊 Summary Report", use_container_width=True):
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
            st.download_button(
                label="⬇️ Download Summary PDF",
                data=pdf_data,
                file_name="Daily_Summary_Report.pdf",
                mime="application/pdf",
                use_container_width=True
            )

