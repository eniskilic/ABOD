import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch, landscape
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
st.set_page_config(page_title="Amazon Blanket Labels – v7.3.1", layout="centered")
st.title("🧵 Amazon Blanket Label Generator — v7.3.1")

st.write("""
### 🪡 Features
- Keeps **original PDF order** (1 label per customization)
- Adds **Quantity** under Order Date  
- Spanish translation (italic) for **Thread Color**
- **Grouped sections** with clean separators  
- Auto-UPPERCASE for Name  
- **Gift Message fix** — now reads only within product borders  
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

            beanie = "Yes" if re.search(r"Personalized Baby Beanie:\s*Yes", block, re.IGNORECASE) else "No"
            gift_box = "Yes" if re.search(r"Gift Box\s*&\s*Gift Card:\s*Yes", block, re.IGNORECASE) else "No"
            gift_note = "Yes" if re.search(r"Gift Message:", block, re.IGNORECASE) else "No"

            # ---- FIXED Gift Message extraction ----
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
    # Generate 6x4 Landscape Labels
    # --------------------------------------
    def draw_separator(c, y_pos, W):
        c.setStrokeColor(colors.lightgrey)
        c.setLineWidth(0.5)
        c.line(0.4 * inch, y_pos, W - 0.4 * inch, y_pos)

    def generate_labels(dataframe):
        buf = BytesIO()
        page_size = landscape((4 * inch, 6 * inch))
        c = canvas.Canvas(buf, pagesize=page_size)
        W, H = page_size
        left = 0.4 * inch
        top = H - 0.4 * inch

        for _, row in dataframe.iterrows():
            y = top

            # --- Top Group ---
            c.setFont("Helvetica-Bold", 11)
            c.drawString(left, y, f"Order ID: {row['Order ID']}")
            y -= 0.22 * inch
            c.setFont("Helvetica", 10)
            c.drawString(left, y, row['Buyer Name'])
            y -= 0.22 * inch
            c.drawString(left, y, f"Order Date: {row['Order Date']}")
            y -= 0.22 * inch
            c.drawString(left, y, f"Quantity: {row['Quantity']}")
            y -= 0.2 * inch
            draw_separator(c, y, W)
            y -= 0.25 * inch

            # --- Product Colors Group ---
            c.setFont("Helvetica-Bold", 13)
            c.drawString(left, y, f"Blanket: {row['Blanket Color']}")
            y -= 0.35 * inch
            c.drawString(left, y, f"Thread: {row['Thread Color']}")
            y -= 0.2 * inch
            draw_separator(c, y, W)
            y -= 0.25 * inch

            # --- Embroidery + Beanie Group ---
            c.setFont("Helvetica-Bold", 15)
            c.drawString(left, y, f"Name: {row['Customization Name']}")
            y -= 0.4 * inch
            c.setFont("Helvetica-Bold", 13)
            c.drawString(left, y, f"Include Beanie: {row['Include Beanie']}")
            y -= 0.2 * inch
            draw_separator(c, y, W)
            y -= 0.25 * inch

            # --- Packaging Group (Gift Box / Note) ---
            c.setFont("Helvetica-Bold", 13)
            c.drawString(left, y, f"Gift Box: {row['Gift Box']}")
            y -= 0.35 * inch
            c.drawString(left, y, f"Gift Note: {row['Gift Note']}")
            y -= 0.2 * inch
            draw_separator(c, y, W)
            y -= 0.2 * inch

            c.showPage()

        c.save()
        buf.seek(0)
        return buf

    # --------------------------------------
    # Download PDF
    # --------------------------------------
    if st.button("📦 Generate Final 6×4 Labels PDF"):
        pdf_data = generate_labels(df)
        st.download_button(
            label="⬇️ Download Clean Labels (v7.3.1)",
            data=pdf_data,
            file_name="Amazon_Labels_v7_3_1.pdf",
            mime="application/pdf"
        )
