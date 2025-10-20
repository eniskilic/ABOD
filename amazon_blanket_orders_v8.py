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
# Airtable Functions (UPDATED)
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
                                # ✅ Buyer Name first
                                "Buyer Name": item_row['Buyer Name'],
                                
                                # ✅ Customization Name now comes right after Order ID
                                "Order ID": [airtable_order_id],
                                "Customization Name": item_row['Customization Name'],
                                
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

    # The rest of your code (summary, PDF generation, download buttons, etc.)
    # remains exactly the same as before.
