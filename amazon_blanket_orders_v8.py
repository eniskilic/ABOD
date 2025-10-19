import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch, landscape, portrait
from reportlab.lib import colors

# --------------------------------------
# Color dictionary for Spanish translation (Thread colors)
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
    """Adds Spanish translation (italic style in label rendering)."""
    if not color:
        return color
    base = color.strip()
    for eng, esp in COLOR_TRANSLATIONS.items():
        if eng.lower() in base.lower():
            return f"{base} ({esp})"
    return base

# --------------------------------------
# Streamlit Setup
# --------------------------------------
st.set_page_config(page_title="Amazon Blanket Labels – v8.0", layout="centered")
st.title("🧵 Amazon Blanket Label Generator — v8.0")

st.write("""
### 🪡 Features
- **Two Label Types**: Manufacturing labels + Gift message labels
- **Two-column layout**: Product info (left) | Packaging info (right)
- **Optimized for B&W printing** with boxes and clear hierarchy
- **Thread Color in bold box** - most prominent
- Spanish translation for Thread Color
- **Embroidered name preserved** as written in PDF
- Font sizes: 14-16pt for easy reading
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

            # Gift Message extraction
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
    df.index = df.index + 1  # Start row numbering from 1 instead of 0
    st.success(f"✅ {len(df)} labels parsed (in original order).")
    st.dataframe(df)

    # --------------------------------------
    # Generate Manufacturing/Packaging Labels (6x4 Landscape) - Two Column Layout
    # --------------------------------------
    def generate_manufacturing_labels(dataframe):
        buf = BytesIO()
        page_size = landscape((4 * inch, 6 * inch))
        c = canvas.Canvas(buf, pagesize=page_size)
        W, H = page_size
        left = 0.3 * inch
        right = W - 0.3 * inch
        top = H - 0.3 * inch
        middle_x = left + (right - left) * 0.60  # 60/40 split

        for _, row in dataframe.iterrows():
            y = top

            # --- Order Info Section ---
            c.setFont("Helvetica-Bold", 14)
            c.drawString(left, y, f"Order ID: {row['Order ID']}")
            c.drawRightString(right, y, f"Qty: {row['Quantity']}")
            y -= 0.25 * inch
            
            c.setFont("Helvetica", 14)
            c.drawString(left, y, f"Buyer: {row['Buyer Name']}")
            c.drawRightString(right, y, f"Date: {row['Order Date']}")
            y -= 0.3 * inch

            # --- COLOR BOX: Thread Color & Blanket Color (Most Important) ---
            box_height = 0.65 * inch
            box_y = y - box_height
            c.setStrokeColor(colors.black)
            c.setLineWidth(2)
            c.rect(left, box_y, right - left, box_height, stroke=1, fill=0)
            
            # Thread Color (top line in box)
            c.setFont("Helvetica-Bold", 16)
            text_y = box_y + box_height - 0.22 * inch
            c.drawString(left + 0.1 * inch, text_y, f"THREAD COLOR: {row['Thread Color']}")
            
            # Blanket Color (bottom line in box)
            text_y -= 0.28 * inch
            c.setFont("Helvetica-Bold", 15)
            c.drawString(left + 0.1 * inch, text_y, f"BLANKET COLOR: {row['Blanket Color']}")
            y = box_y - 0.3 * inch

            # --- Vertical Divider Line ---
            divider_start_y = y
            c.setStrokeColor(colors.grey)
            c.setLineWidth(1)
            c.line(middle_x, y, middle_x, 0.3 * inch)

            # --- LEFT COLUMN: PRODUCT INFO ---
            y_left = y
            c.setFont("Helvetica-Bold", 15)
            c.drawString(left, y_left, "PRODUCT:")
            y_left -= 0.3 * inch

            # Name with wrapping if too long
            c.setFont("Helvetica-Bold", 15)
            name_text = f"★ Name: {row['Customization Name']}"
            left_column_width = middle_x - left - 0.2 * inch  # Available width for left column
            
            # Check if name fits in one line
            if c.stringWidth(name_text, "Helvetica-Bold", 15) <= left_column_width:
                # Fits in one line
                c.drawString(left, y_left, name_text)
                y_left -= 0.3 * inch
            else:
                # Need to wrap - split into two lines
                c.drawString(left, y_left, "★ Name:")
                y_left -= 0.25 * inch
                # Draw name on second line with indent
                name_only = row['Customization Name']
                
                # If still too long, try to split into words
                words = name_only.split()
                if len(words) > 1 and c.stringWidth(name_only, "Helvetica-Bold", 15) > left_column_width - 0.3 * inch:
                    # Split into multiple lines if needed
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

            # --- RIGHT COLUMN: PACKAGING INFO ---
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

    # --------------------------------------
    # Generate Gift Message Labels (6x4 Landscape - Colorful & Centered)
    # --------------------------------------
    def generate_gift_message_labels(dataframe):
        buf = BytesIO()
        page_size = landscape((4 * inch, 6 * inch))
        c = canvas.Canvas(buf, pagesize=page_size)
        W, H = page_size

        # Filter only orders with gift messages
        gift_orders = dataframe[dataframe['Gift Message'] != ""]

        if len(gift_orders) == 0:
            # Create a blank page with message
            c.setFont("Helvetica", 14)
            c.drawCentredString(W / 2, H / 2, "No gift messages found in orders")
            c.showPage()
        else:
            for _, row in gift_orders.iterrows():
                # --- White Background ---
                c.setFillColor(colors.white)
                c.rect(0, 0, W, H, fill=1, stroke=0)
                
                # --- Single Solid Border Frame ---
                c.setStrokeColor(colors.HexColor("#FFB6C1"))  # Light pink border
                c.setLineWidth(3)  # 3pt thickness
                c.setDash([])  # Ensure solid line, not dashed
                c.rect(0.4 * inch, 0.4 * inch, W - 0.8 * inch, H - 0.8 * inch, stroke=1, fill=0)

                # --- Message Text (Centered & Colorful) ---
                c.setFillColor(colors.HexColor("#4A4A4A"))  # Dark gray for text
                c.setFont("Helvetica", 16)
                message = row['Gift Message']
                
                # Simple text wrapping
                words = message.split()
                lines = []
                current_line = []
                max_width = W - 1.2 * inch  # Leave margins
                
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

                # Calculate starting y position to vertically center all lines
                total_height = len(lines) * 0.28 * inch
                y = (H + total_height) / 2

                # Draw centered lines
                for line in lines:
                    c.drawCentredString(W / 2, y, line)
                    y -= 0.28 * inch

                c.showPage()

        c.save()
        buf.seek(0)
        return buf

    # --------------------------------------
    # Download Buttons
    # --------------------------------------
    st.write("---")
    st.subheader("📥 Generate Your Labels")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📦 Manufacturing Labels", use_container_width=True):
            pdf_data = generate_manufacturing_labels(df)
            st.download_button(
                label="⬇️ Download Manufacturing Labels",
                data=pdf_data,
                file_name="Manufacturing_Labels_v8.pdf",
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
                file_name="Gift_Message_Labels_v8.pdf",
                mime="application/pdf",
                use_container_width=True
            )
