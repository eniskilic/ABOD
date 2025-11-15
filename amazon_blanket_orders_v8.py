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
# Dark Mode Custom CSS Styling
# --------------------------------------
st.markdown("""
<style>
    /* Dark Mode Base */
    .main {
        background: #0f1419;
        color: #e4e6eb;
    }
    
    .stApp {
        background: #0f1419;
    }
    
    /* Sidebar Dark Styling */
    [data-testid="stSidebar"] {
        background: #1a1f2e;
        border-right: 1px solid #2d3748;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: #e4e6eb;
    }
    
    /* Sidebar Navigation Links */
    .nav-link {
        display: block;
        padding: 12px 15px;
        margin: 4px 0;
        border-radius: 10px;
        color: #a0aec0;
        text-decoration: none;
        transition: all 0.2s ease;
        cursor: pointer;
    }
    
    .nav-link:hover {
        background: #2d3748;
        color: #e4e6eb;
        text-decoration: none;
    }
    
    .nav-link.active {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
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
        transition: all 0.3s ease;
        border-color: #667eea;
    }
    
    [data-testid="stMetric"] label {
        font-size: 0.85em !important;
        color: #a0aec0 !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 2.5em !important;
        font-weight: 700 !important;
        color: #e4e6eb !important;
    }
    
    /* Headers Dark */
    h1 {
        color: #e4e6eb;
        font-weight: 700;
        padding-bottom: 15px;
        border-bottom: 3px solid #667eea;
        margin-bottom: 30px;
    }
    
    h2 {
        color: #e4e6eb;
        font-weight: 600;
        margin-top: 40px;
        margin-bottom: 20px;
    }
    
    h3 {
        color: #cbd5e0;
        font-weight: 600;
        margin-bottom: 15px;
    }
    
    /* Buttons Dark */
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
    
    /* File Uploader Dark */
    [data-testid="stFileUploader"] {
        background: #1a1f2e;
        padding: 40px;
        border-radius: 12px;
        border: 2px dashed #2d3748;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #667eea;
        background: #1e2432;
    }
    
    [data-testid="stFileUploader"] label {
        color: #e4e6eb !important;
    }
    
    [data-testid="stFileUploader"] section {
        border-color: #2d3748 !important;
    }
    
    /* Info boxes Dark */
    .stAlert {
        background: linear-gradient(135deg, #667eea20, #764ba220) !important;
        border: 1px solid #667eea40 !important;
        border-radius: 10px;
        border-left: 4px solid #667eea !important;
        color: #cbd5e0 !important;
    }
    
    /* Success boxes */
    [data-baseweb="notification"] {
        background: #1a1f2e !important;
        border: 1px solid #48bb78 !important;
        color: #e4e6eb !important;
    }
    
    /* Dataframe Dark */
    [data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
    }
    
    [data-testid="stDataFrame"] table {
        background: #1a1f2e !important;
        color: #e4e6eb !important;
    }
    
    [data-testid="stDataFrame"] thead tr th {
        background: #2d3748 !important;
        color: #e4e6eb !important;
    }
    
    [data-testid="stDataFrame"] tbody tr {
        background: #1e2432 !important;
        color: #cbd5e0 !important;
    }
    
    [data-testid="stDataFrame"] tbody tr:hover {
        background: #252d3d !important;
    }
    
    /* Expander Dark */
    [data-testid="stExpander"] {
        background: #1a1f2e !important;
        border-radius: 10px;
        border: 1px solid #2d3748 !important;
        margin-bottom: 10px;
    }
    
    [data-testid="stExpander"] [data-testid="stMarkdownContainer"] {
        color: #e4e6eb !important;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Download button */
    .stDownloadButton button {
        background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
        width: 100%;
    }
    
    .stDownloadButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 5px 15px rgba(72, 187, 120, 0.4);
    }
    
    /* Text color overrides */
    p, span, div {
        color: #cbd5e0;
    }
    
    strong {
        color: #e4e6eb;
    }
    
    /* Section divider */
    hr {
        border: none;
        border-top: 1px solid #2d3748;
        margin: 40px 0;
    }
    
    /* Spinner Dark */
    .stSpinner > div {
        border-top-color: #667eea !important;
    }
    
    /* Input fields */
    input, textarea, select {
        background: #1a1f2e !important;
        color: #e4e6eb !important;
        border: 1px solid #2d3748 !important;
    }
    
    /* Markdown text */
    .stMarkdown {
        color: #cbd5e0 !important;
    }
    
    /* Status indicator */
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
# CRITICAL FIX: Safe Label Merging with Validation
# --------------------------------------

def extract_order_ids_from_shipping_pdf(shipping_pdf_bytes):
    """
    Extract Order IDs from shipping label PDF.
    Amazon shipping labels contain Order IDs in format: XXX-XXXXXXX-XXXXXXX

    Returns:
        list: List of (page_index, order_id) tuples
    """
    extracted_orders = []
    order_id_pattern = re.compile(r'\b(\d{3}-\d{7}-\d{7})\b')

    try:
        # Use pdfplumber to extract text from each page
        with pdfplumber.open(shipping_pdf_bytes) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                text = page.extract_text() or ""

                # Search for Order ID pattern
                matches = order_id_pattern.findall(text)

                if matches:
                    # Take the first match (should be the order ID for this label)
                    order_id = matches[0]
                    extracted_orders.append((page_idx, order_id))
                else:
                    # No Order ID found on this page
                    extracted_orders.append((page_idx, None))

        return extracted_orders

    except Exception as e:
        st.error(f"Error extracting Order IDs from shipping labels: {str(e)}")
        return []


def validate_label_mapping(shipping_pdf_bytes, order_dataframe):
    """
    Validate that shipping labels match the orders in the dataframe.

    Returns:
        dict: Validation report with structure:
        {
            'is_valid': bool,
            'shipping_labels': list of (page_idx, order_id),
            'expected_orders': list of order_ids from dataframe,
            'matched_orders': list of order_ids that have shipping labels,
            'missing_labels': list of order_ids without shipping labels,
            'extra_labels': list of shipping label order_ids not in dataframe,
            'unreadable_labels': list of page indices where Order ID couldn't be extracted,
            'mapping': dict mapping order_id -> (shipping_page_idx, mfg_page_indices),
            'warnings': list of warning messages,
            'errors': list of error messages
        }
    """
    validation_report = {
        'is_valid': False,
        'shipping_labels': [],
        'expected_orders': [],
        'matched_orders': [],
        'missing_labels': [],
        'extra_labels': [],
        'unreadable_labels': [],
        'mapping': {},
        'warnings': [],
        'errors': []
    }

    try:
        # Step 1: Extract Order IDs from shipping labels
        shipping_labels = extract_order_ids_from_shipping_pdf(shipping_pdf_bytes)
        validation_report['shipping_labels'] = shipping_labels

        # Track unreadable labels
        for page_idx, order_id in shipping_labels:
            if order_id is None:
                validation_report['unreadable_labels'].append(page_idx)
                validation_report['warnings'].append(
                    f"⚠️ Shipping label page {page_idx + 1}: Could not extract Order ID"
                )

        # Step 2: Get expected order sequence from dataframe
        seen_orders = []
        order_item_counts = {}

        for order_id in order_dataframe['Order ID']:
            if order_id not in seen_orders:
                seen_orders.append(order_id)
                order_item_counts[order_id] = len(order_dataframe[order_dataframe['Order ID'] == order_id])

        validation_report['expected_orders'] = seen_orders

        # Step 3: Create mapping of shipping labels to Order IDs
        shipping_order_map = {}
        for page_idx, order_id in shipping_labels:
            if order_id:
                shipping_order_map[order_id] = page_idx

        # Step 4: Compare shipping labels with expected orders
        for order_id in seen_orders:
            if order_id in shipping_order_map:
                validation_report['matched_orders'].append(order_id)
            else:
                validation_report['missing_labels'].append(order_id)
                validation_report['errors'].append(
                    f"❌ Order {order_id}: No shipping label found"
                )

        # Check for extra shipping labels not in dataframe
        for order_id, page_idx in shipping_order_map.items():
            if order_id not in seen_orders:
                validation_report['extra_labels'].append(order_id)
                validation_report['warnings'].append(
                    f"⚠️ Shipping label page {page_idx + 1} ({order_id}): Not found in manufacturing orders"
                )

        # Step 5: Build safe mapping for merge
        # Only map orders that have matching shipping labels
        mfg_index = 0
        for expected_order_id in seen_orders:
            if expected_order_id in shipping_order_map:
                shipping_page = shipping_order_map[expected_order_id]
                item_count = order_item_counts[expected_order_id]
                mfg_pages = list(range(mfg_index, mfg_index + item_count))

                validation_report['mapping'][expected_order_id] = {
                    'shipping_page': shipping_page,
                    'mfg_pages': mfg_pages,
                    'item_count': item_count
                }

            # Always increment mfg_index to maintain alignment
            mfg_index += order_item_counts[expected_order_id]

        # Step 6: Determine if validation passed
        validation_report['is_valid'] = (
            len(validation_report['missing_labels']) == 0 and
            len(validation_report['unreadable_labels']) == 0
        )

        # Summary messages
        if validation_report['is_valid']:
            validation_report['summary'] = f"✅ Perfect match: All {len(seen_orders)} orders have shipping labels"
        else:
            missing_count = len(validation_report['missing_labels'])
            unreadable_count = len(validation_report['unreadable_labels'])
            validation_report['summary'] = f"❌ Validation failed: {missing_count} missing labels, {unreadable_count} unreadable labels"

        return validation_report

    except Exception as e:
        validation_report['errors'].append(f"Critical error during validation: {str(e)}")
        return validation_report


def merge_shipping_and_manufacturing_labels_safe(
    shipping_pdf_bytes,
    manufacturing_pdf_bytes,
    validation_report,
    mode='strict'
):
    """
    Safely merge shipping labels with manufacturing labels using validated mapping.

    Args:
        shipping_pdf_bytes: BytesIO object containing shipping labels PDF
        manufacturing_pdf_bytes: BytesIO object containing manufacturing labels PDF
        validation_report: Validation report from validate_label_mapping()
        mode: 'strict' (only merge validated labels) or 'flexible' (allow partial merges)

    Returns:
        tuple: (merged_pdf_buffer, num_shipping_labels, num_manufacturing_labels, merge_details)
    """
    try:
        if mode == 'strict' and not validation_report['is_valid']:
            return None, 0, 0, {
                'success': False,
                'message': 'Validation failed. Cannot merge in STRICT mode.',
                'details': validation_report['errors']
            }

        shipping_pdf = PdfReader(shipping_pdf_bytes)
        manufacturing_pdf = PdfReader(manufacturing_pdf_bytes)

        output_pdf = PdfWriter()
        merged_count = 0
        mfg_labels_added = 0
        merge_log = []

        # Use the validated mapping
        for order_id, mapping_info in validation_report['mapping'].items():
            shipping_page = mapping_info['shipping_page']
            mfg_pages = mapping_info['mfg_pages']
            item_count = mapping_info['item_count']

            # Verify shipping page exists
            if shipping_page >= len(shipping_pdf.pages):
                merge_log.append(f"⚠️ Skipped {order_id}: Shipping page {shipping_page} out of range")
                continue

            # Add shipping label
            output_pdf.add_page(shipping_pdf.pages[shipping_page])
            merged_count += 1

            # Add corresponding manufacturing labels
            labels_for_this_order = 0
            for mfg_idx in mfg_pages:
                if mfg_idx < len(manufacturing_pdf.pages):
                    output_pdf.add_page(manufacturing_pdf.pages[mfg_idx])
                    mfg_labels_added += 1
                    labels_for_this_order += 1
                else:
                    merge_log.append(f"⚠️ {order_id}: Manufacturing page {mfg_idx} out of range")

            merge_log.append(
                f"✅ {order_id}: 1 shipping label + {labels_for_this_order} manufacturing label(s)"
            )

        # Write output
        output_buffer = BytesIO()
        output_pdf.write(output_buffer)
        output_buffer.seek(0)

        merge_details = {
            'success': True,
            'message': f'Successfully merged {merged_count} orders',
            'merged_count': merged_count,
            'mfg_labels_added': mfg_labels_added,
            'log': merge_log,
            'skipped_orders': validation_report['missing_labels']
        }

        return output_buffer, merged_count, mfg_labels_added, merge_details

    except Exception as e:
        return None, 0, 0, {
            'success': False,
            'message': f'Error during merge: {str(e)}',
            'details': []
        }


# --------------------------------------
# LEGACY: Old merge function (kept for compatibility)
# --------------------------------------
def merge_shipping_and_manufacturing_labels(shipping_pdf_bytes, manufacturing_pdf_bytes, order_dataframe):
    """
    DEPRECATED: Use merge_shipping_and_manufacturing_labels_safe() instead.
    This function assumes perfect 1:1 mapping and does NOT validate Order IDs.
    """
    try:
        shipping_pdf = PdfReader(shipping_pdf_bytes)
        manufacturing_pdf = PdfReader(manufacturing_pdf_bytes)

        seen_orders = []
        order_item_counts = []

        for order_id in order_dataframe['Order ID']:
            if order_id not in seen_orders:
                seen_orders.append(order_id)
                item_count = len(order_dataframe[order_dataframe['Order ID'] == order_id])
                order_item_counts.append(item_count)

        shipping_to_mfg = {}
        mfg_index = 0

        for shipping_index, item_count in enumerate(order_item_counts):
            shipping_to_mfg[shipping_index] = list(range(mfg_index, mfg_index + item_count))
            mfg_index += item_count

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
# PDF Generation Functions (same as before)
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
# SIDEBAR WITH FUNCTIONAL NAVIGATION
# --------------------------------------
with st.sidebar:
    st.markdown("# 🧵 Blanket Manager")
    st.markdown("### Version 11.0 - Safe Merge")
    st.markdown("---")
    
    st.markdown("#### 📋 Quick Navigation")
    
    # Functional navigation links with anchor tags
    st.markdown('<a href="#upload-order" class="nav-link">📄 Upload Order</a>', unsafe_allow_html=True)
    st.markdown('<a href="#dashboard" class="nav-link">📊 Dashboard</a>', unsafe_allow_html=True)
    st.markdown('<a href="#color-analytics" class="nav-link">🎨 Color Analytics</a>', unsafe_allow_html=True)
    st.markdown('<a href="#bobbin-setup" class="nav-link">🧵 Bobbin Setup</a>', unsafe_allow_html=True)
    st.markdown('<a href="#generate-labels" class="nav-link">📥 Generate Labels</a>', unsafe_allow_html=True)
    st.markdown('<a href="#label-merge" class="nav-link">🔄 Label Merge</a>', unsafe_allow_html=True)
    st.markdown('<a href="#airtable-sync" class="nav-link">☁️ Airtable Sync</a>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("#### ✨ Features")
    st.markdown("✓ PDF Parsing")
    st.markdown("✓ Label Generation")
    st.markdown("✓ Order Merging")
    st.markdown("✓ Cloud Sync")
    st.markdown("✓ Spanish Translation")
    st.markdown("✓ Duplicate Detection")
    
    st.markdown("---")
    st.markdown('<div class="status-indicator"><div class="status-dot"></div><span>System Ready</span></div>', unsafe_allow_html=True)

# --------------------------------------
# MAIN CONTENT
# --------------------------------------
st.title("🧵 Amazon Blanket Order Manager")

st.markdown("""
**Professional order processing & label generation system**  
Parse Amazon PDFs • Generate labels • Merge shipments • Sync to Airtable
""")

st.markdown("---")

# File Upload Section with anchor
st.markdown('<a id="upload-order"></a>', unsafe_allow_html=True)
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
    # Dashboard Metrics with anchor
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="dashboard"></a>', unsafe_allow_html=True)
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
    # Color Breakdown with anchor
    # --------------------------------------
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
    
    # --------------------------------------
    # Bobbin Setup Section with anchor
    # --------------------------------------
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

    # --------------------------------------
    # Generate Labels Section with anchor
    # --------------------------------------
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
    # Label Merging Section with anchor - UPDATED WITH VALIDATION
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="label-merge"></a>', unsafe_allow_html=True)
    st.markdown("## 🔄 Merge Shipping & Manufacturing Labels")

    st.info("""
    **🛡️ NEW: Safe Label Merging with Validation**
    1. Generate Manufacturing Labels above (click the button)
    2. Upload your shipping labels PDF from Amazon/UPS
    3. **Validate** - System will extract Order IDs from shipping labels and verify alignment
    4. **Review** - Check which orders have shipping labels before merging
    5. **Merge** - Create combined PDF only for validated orders

    ✨ **Version 11.0:** Now validates Order IDs to prevent misaligned labels!
    """)

    shipping_labels_upload = st.file_uploader(
        "📤 Upload Shipping Labels PDF",
        type=["pdf"],
        key="shipping_labels",
        help="Upload the shipping labels PDF from Amazon or your carrier"
    )

    # Initialize session state for validation
    if 'validation_report' not in st.session_state:
        st.session_state.validation_report = None
    if 'validation_done' not in st.session_state:
        st.session_state.validation_done = False

    if shipping_labels_upload and st.session_state.manufacturing_labels_buffer:

        # STEP 1: VALIDATION
        st.markdown("### Step 1: Validate Label Alignment")

        col_validate1, col_validate2 = st.columns([2, 1])

        with col_validate1:
            if st.button("🔍 Validate Shipping Labels", type="primary", use_container_width=True):
                with st.spinner("🔍 Extracting Order IDs from shipping labels and validating..."):
                    shipping_labels_upload.seek(0)
                    validation_report = validate_label_mapping(shipping_labels_upload, df)
                    st.session_state.validation_report = validation_report
                    st.session_state.validation_done = True

        with col_validate2:
            st.metric("Expected Orders", df['Order ID'].nunique())
            st.metric("Total Items", len(df))

        # STEP 2: SHOW VALIDATION RESULTS
        if st.session_state.validation_done and st.session_state.validation_report:
            validation_report = st.session_state.validation_report

            st.markdown("---")
            st.markdown("### Step 2: Validation Results")

            # Display summary
            if validation_report['is_valid']:
                st.success(validation_report['summary'])
            else:
                st.error(validation_report['summary'])

            # Metrics
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)

            with col_m1:
                st.metric("✅ Matched Orders", len(validation_report['matched_orders']))
            with col_m2:
                st.metric("❌ Missing Labels", len(validation_report['missing_labels']))
            with col_m3:
                st.metric("⚠️ Unreadable Labels", len(validation_report['unreadable_labels']))
            with col_m4:
                st.metric("📄 Extra Labels", len(validation_report['extra_labels']))

            # Detailed breakdown
            with st.expander("📋 Detailed Validation Report", expanded=not validation_report['is_valid']):

                # Orders with shipping labels
                if validation_report['matched_orders']:
                    st.markdown("#### ✅ Orders with Shipping Labels")
                    matched_df_data = []
                    for order_id in validation_report['matched_orders']:
                        buyer = df[df['Order ID'] == order_id]['Buyer Name'].iloc[0]
                        item_count = len(df[df['Order ID'] == order_id])
                        shipping_page = validation_report['mapping'][order_id]['shipping_page']
                        matched_df_data.append({
                            'Order ID': order_id,
                            'Buyer Name': buyer,
                            'Items': item_count,
                            'Shipping Label Page': shipping_page + 1
                        })
                    st.dataframe(pd.DataFrame(matched_df_data), use_container_width=True)

                # Orders missing shipping labels
                if validation_report['missing_labels']:
                    st.markdown("#### ❌ Orders WITHOUT Shipping Labels")
                    st.warning("⚠️ **CRITICAL:** These orders will NOT be included in the merge!")
                    missing_df_data = []
                    for order_id in validation_report['missing_labels']:
                        buyer = df[df['Order ID'] == order_id]['Buyer Name'].iloc[0]
                        item_count = len(df[df['Order ID'] == order_id])
                        missing_df_data.append({
                            'Order ID': order_id,
                            'Buyer Name': buyer,
                            'Items': item_count,
                            'Status': '⚠️ No shipping label found'
                        })
                    st.dataframe(pd.DataFrame(missing_df_data), use_container_width=True)

                # Unreadable labels
                if validation_report['unreadable_labels']:
                    st.markdown("#### ⚠️ Unreadable Shipping Labels")
                    st.warning("Could not extract Order ID from these shipping label pages:")
                    for page_idx in validation_report['unreadable_labels']:
                        st.write(f"• Page {page_idx + 1}: Order ID not found in PDF text")

                # Extra labels
                if validation_report['extra_labels']:
                    st.markdown("#### 📄 Extra Shipping Labels")
                    st.info("These shipping labels are not in your manufacturing orders:")
                    for order_id in validation_report['extra_labels']:
                        st.write(f"• Order ID: {order_id}")

                # Show all warnings and errors
                if validation_report['warnings']:
                    st.markdown("#### ⚠️ Warnings")
                    for warning in validation_report['warnings']:
                        st.warning(warning)

                if validation_report['errors']:
                    st.markdown("#### ❌ Errors")
                    for error in validation_report['errors']:
                        st.error(error)

            # STEP 3: MERGE MODE SELECTION
            st.markdown("---")
            st.markdown("### Step 3: Choose Merge Mode")

            col_mode1, col_mode2 = st.columns(2)

            with col_mode1:
                st.markdown("#### 🛡️ STRICT Mode (Recommended)")
                st.markdown("""
                - ✅ Only merge if ALL orders have shipping labels
                - ✅ Prevents any misalignment
                - ❌ Blocks merge if validation fails
                """)
                if validation_report['is_valid']:
                    strict_enabled = True
                    st.success("✅ STRICT mode available")
                else:
                    strict_enabled = False
                    st.error("❌ Cannot use STRICT mode - validation failed")

            with col_mode2:
                st.markdown("#### ⚙️ FLEXIBLE Mode")
                st.markdown("""
                - ✅ Merge only orders that have shipping labels
                - ⚠️ Skip orders without shipping labels
                - ℹ️ Use when some orders are cancelled/delayed
                """)
                if len(validation_report['matched_orders']) > 0:
                    flexible_enabled = True
                    st.info(f"ℹ️ Will merge {len(validation_report['matched_orders'])} orders")
                else:
                    flexible_enabled = False
                    st.error("❌ No orders to merge")

            # STEP 4: MERGE WITH CONFIRMATION
            st.markdown("---")
            st.markdown("### Step 4: Review & Merge")

            merge_mode = st.radio(
                "Select merge mode:",
                options=["strict", "flexible"],
                format_func=lambda x: "🛡️ STRICT - All orders must have labels" if x == "strict" else "⚙️ FLEXIBLE - Merge available orders only",
                disabled=not (strict_enabled or flexible_enabled),
                index=0 if strict_enabled else (1 if flexible_enabled else 0)
            )

            # Merge preview
            with st.expander("📋 Merge Preview - What will be merged?", expanded=True):
                preview_data = []
                for order_id, mapping_info in validation_report['mapping'].items():
                    buyer = df[df['Order ID'] == order_id]['Buyer Name'].iloc[0]
                    preview_data.append({
                        'Order ID': order_id,
                        'Buyer': buyer,
                        'Shipping Label Page': mapping_info['shipping_page'] + 1,
                        'Manufacturing Labels': f"{mapping_info['item_count']} label(s)",
                        'Status': '✅ Ready to merge'
                    })

                if preview_data:
                    st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
                    st.success(f"**{len(preview_data)} order(s) will be merged** ({sum([m['item_count'] for m in validation_report['mapping'].values()])} manufacturing labels total)")
                else:
                    st.error("No orders available for merging")

            # Merge button
            can_merge = (merge_mode == "strict" and strict_enabled) or (merge_mode == "flexible" and flexible_enabled)

            if can_merge:
                col_merge_btn1, col_merge_btn2 = st.columns([3, 1])

                with col_merge_btn1:
                    if st.button("✅ CONFIRM AND MERGE LABELS", type="primary", use_container_width=True):
                        with st.spinner(f"Merging labels in {merge_mode.upper()} mode..."):
                            shipping_labels_upload.seek(0)
                            st.session_state.manufacturing_labels_buffer.seek(0)

                            merged_pdf, num_shipping, num_manufacturing, merge_details = merge_shipping_and_manufacturing_labels_safe(
                                shipping_labels_upload,
                                st.session_state.manufacturing_labels_buffer,
                                validation_report,
                                mode=merge_mode
                            )

                            if merged_pdf and merge_details['success']:
                                st.success(f"🎉 {merge_details['message']}")

                                # Show merge details
                                with st.expander("📊 Merge Details", expanded=True):
                                    col_d1, col_d2 = st.columns(2)
                                    with col_d1:
                                        st.metric("Orders Merged", merge_details['merged_count'])
                                    with col_d2:
                                        st.metric("Manufacturing Labels", merge_details['mfg_labels_added'])

                                    if merge_details['log']:
                                        st.markdown("#### Merge Log")
                                        for log_entry in merge_details['log']:
                                            if '✅' in log_entry:
                                                st.success(log_entry)
                                            elif '⚠️' in log_entry:
                                                st.warning(log_entry)

                                    if merge_details['skipped_orders']:
                                        st.warning(f"⚠️ Skipped {len(merge_details['skipped_orders'])} order(s) without shipping labels:")
                                        for skipped_order in merge_details['skipped_orders']:
                                            buyer = df[df['Order ID'] == skipped_order]['Buyer Name'].iloc[0]
                                            st.write(f"• {skipped_order} - {buyer}")

                                st.download_button(
                                    label="⬇️ Download Merged Labels PDF",
                                    data=merged_pdf,
                                    file_name="Merged_Shipping_Manufacturing_Labels_VALIDATED.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                            else:
                                st.error(f"❌ Merge failed: {merge_details['message']}")
                                if merge_details.get('details'):
                                    with st.expander("Error Details"):
                                        for detail in merge_details['details']:
                                            st.write(detail)

                with col_merge_btn2:
                    if st.button("🔄 Start Over", use_container_width=True):
                        st.session_state.validation_done = False
                        st.session_state.validation_report = None
                        st.rerun()
            else:
                st.error("Cannot merge: Validation failed and no valid orders available")
                st.info("💡 Tip: Check with Amazon why some orders don't have shipping labels. They may be cancelled or on hold.")

    elif shipping_labels_upload and not st.session_state.manufacturing_labels_buffer:
        st.warning("⚠️ Please generate Manufacturing Labels first (click the button above)")

    elif not shipping_labels_upload and st.session_state.manufacturing_labels_buffer:
        st.info("📤 Upload your shipping labels PDF above to begin validation")

    # --------------------------------------
    # Airtable Upload Section with anchor
    # --------------------------------------
    st.markdown("---")
    st.markdown('<a id="airtable-sync"></a>', unsafe_allow_html=True)
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
<div style='text-align: center; color: #a0aec0; padding: 20px;'>
    <p><strong>Amazon Blanket Order Manager v11.0 - Safe Merge</strong></p>
    <p>Professional order processing & label generation system with validated merging</p>
</div>
""", unsafe_allow_html=True)

