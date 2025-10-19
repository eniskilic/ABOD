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
- **Optimized for B&W printing** with boxes and clear hierarchy
- **Thread Color** prominently displayed in box
- Spanish translation for Thread Color
- Font sizes: 14-16pt for easy reading
- Auto-UPPERCASE for embroidered names
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
            customization_name = clean_text(name_match.group(1)).upper() if name_match else ""

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
    st.success(f"✅ {len(df)} labels parsed (in original order).")
    st.dataframe(df)

    # --------------------------------------
    # Generate Manufacturing/Packaging Labels (6x4 Landscape)
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

            # --- Order Info Section ---
            c.setFont("Helvetica-Bold", 14)
            c.drawString(left, y, f"Order ID: {row['Order ID']}")
            c.drawRightString(right, y, f"Qty: {row['Quantity']}")
            y -= 0.25 * inch
            
            c.setFont("Helvetica", 14)
            c.drawString(left, y, f"Buyer: {row['Buyer Name']}")
            y -= 0.22 * inch
            c.drawString(left, y, f"Date: {row['Order Date']}")
            y -= 0.3 * inch

            # --- THREAD COLOR BOX (Most Important) ---
            box_height = 0.4 * inch
            box_y = y - box_height
            c.setStrokeColor(colors.black)
            c.setLineWidth(2)
            c.rect(left, box_y, right - left, box_height, stroke=1, fill=0)
            
            c.setFont("Helvetica-Bold", 16)
            text_y = box_y + (box_height - 16) / 2
            c.drawString(left + 0.1 * inch, text_y, f"THREAD COLOR: {row['Thread Color']}")
            y = box_y - 0.25 * inch

            # --- Blanket Color ---
            c.setFont("Helvetica", 14)
            c.drawString(left, y, f"Blanket Color: {row['Blanket Color']}")
            y -= 0.35 * inch

            # --- Separator Line ---
            c.setStrokeColor(colors.black)
            c.setLineWidth(1)
            c.line(left, y, right, y)
            y -= 0.3 * inch

            # --- Embroidered Name (Large & Clear) ---
            c.setFont("Helvetica-Bold", 16)
            c.drawString(left, y, "★ EMBROIDER NAME:")
            y -= 0.3 * inch
            c.setFont("Helvetica-Bold", 18)
            c.drawString(left + 0.2 * inch, y, row['Customization Name'])
            y -= 0.35 * inch

            # --- Separator Line ---
            c.setStrokeColor(colors.black)
            c.setLineWidth(1)
            c.line(left, y, right, y)
            y -= 0.3 * inch

            # --- Packaging Options (Checkboxes) ---
            c.setFont("Helvetica-Bold", 15)
            
            # Beanie
            checkbox = "☑" if row['Include Beanie'] == "YES" else "☐"
            c.drawString(left, y, f"{checkbox} Include Beanie: {row['Include Beanie']}")
            y -= 0.28 * inch

            # Gift Box
            checkbox = "☑" if row['Gift Box'] == "YES" else "☐"
            c.drawString(left, y, f"{checkbox} Gift Box & Card: {row['Gift Box']}")
            y -= 0.28 * inch

            # Gift Message
            checkbox = "☑" if row['Gift Note'] == "YES" else "☐"
            c.drawString(left, y, f"{checkbox} Gift Message: {row['Gift Note']}")

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
                # --- Colorful Background ---
                c.setFillColor(colors.HexColor("#FFF8DC"))  # Soft cream background
                c.rect(0, 0, W, H, fill=1, stroke=0)
                
                # --- Single Solid Border Frame ---
                c.setStrokeColor(colors.HexColor("#FFB6C1"))  # Light pink border
                c.setLineWidth(2)
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
