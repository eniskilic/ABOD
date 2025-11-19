"""
Universal Label Merger with Improved Parser
Successfully tested on UPS, USPS, and FedEx label formats
"""

import pdfplumber
import re
from pypdf import PdfWriter, PdfReader
import pandas as pd
from rapidfuzz import fuzz
import io
from typing import Dict, List, Tuple, Optional
import streamlit as st

class ImprovedLabelMerger:
    """
    Label merger that correctly extracts names from all carrier formats
    """
    
    def extract_name_from_shipping_label(self, text: str) -> Optional[str]:
        """
        Extract recipient name from shipping label - tested on real samples
        Works with: UPS Ground, UPS Ground Saver, USPS Priority Mail, 
        USPS Ground Advantage, FedEx Express, and more
        """
        lines_raw = text.split('\n')
        
        # Strategy 1: Look for "SHIP TO:" with name on next line (UPS format)
        for i, line in enumerate(lines_raw):
            if 'SHIP TO:' in line.upper():
                # Skip empty lines after SHIP TO:
                j = i + 1
                while j < len(lines_raw) and not lines_raw[j].strip():
                    j += 1
                if j < len(lines_raw):
                    name = lines_raw[j].strip()
                    if self._is_valid_name(name):
                        return self._clean_name(name)
        
        # Strategy 2: USPS format - SHIP and TO: on separate or same lines
        for i in range(len(lines_raw) - 1):
            current = lines_raw[i].strip()
            next_line = lines_raw[i + 1].strip()
            
            # Pattern 1: SHIP and TO: on consecutive lines
            if 'SHIP' in current.upper():
                if next_line.upper().startswith('TO:'):
                    # Extract name from TO: line or the line after
                    if ':' in next_line:
                        parts = next_line.split(':', 1)
                        name = parts[1].strip()
                        # USPS often has name on same line as TO:
                        if name and self._is_valid_name(name):
                            return self._clean_name(name)
                    # Check line after TO:
                    if i + 2 < len(lines_raw):
                        name = lines_raw[i + 2].strip()
                        if self._is_valid_name(name):
                            return self._clean_name(name)
            
            # Pattern 2: SHIP TO: on same line (for USPS Priority/Ground)
            if current.upper().startswith('SHIP'):
                # Check if name is on the rest of this line
                if len(current) > 4:
                    name = current[4:].strip()
                    if self._is_valid_name(name):
                        return self._clean_name(name)
        
        # Strategy 3: FedEx format - standalone "TO" with name on next line  
        for i, line in enumerate(lines_raw):
            line = line.strip()
            if line.upper().startswith('TO ') or line.upper() == 'TO':
                # For FedEx, name might be on same line after TO
                if len(line) > 3 and line[2:].strip():
                    name = line[2:].strip()
                    if self._is_valid_name(name):
                        return self._clean_name(name)
                # Or on next non-empty line
                j = i + 1
                while j < len(lines_raw):
                    next_line = lines_raw[j].strip()
                    if next_line:
                        if self._is_valid_name(next_line):
                            return self._clean_name(next_line)
                        break
                    j += 1
        
        return None
    
    def _is_valid_name(self, text: str) -> bool:
        """Check if text could be a person's name"""
        if not text or len(text) < 3:
            return False
        
        # Must have letters
        if not re.search(r'[A-Za-z]', text):
            return False
        
        # Exclude if it's an address (has numbers at start)
        if re.match(r'^\d+\s+', text):
            return False
        
        # Exclude if it ends with a ZIP code pattern
        if re.search(r'\b\d{5}(?:-\d{4})?\s*$', text):
            return False
        
        # Exclude if it looks like a state + zip (e.g., "CA 93704-2145")
        if re.match(r'^[A-Z]{2}\s+\d{5}', text.upper()):
            return False
        
        # Exclude common non-name terms and street suffixes
        exclude_terms = ['TRACKING', 'USPS', 'UPS', 'FEDEX', 'GROUND', 'PRIORITY', 
                        'ADVANTAGE', 'EXPRESS', 'MAIL', 'UNITED STATES',
                        'POSTAGE', 'LBS', 'OZ', 'DWT']
        
        street_suffixes = [' AVE', ' ST', ' RD', ' DR', ' BLVD', ' PL', ' WAY', 
                          ' CT', ' CIR', ' LN', ' PKWY', ' TER', ' TRAIL']
        
        text_upper = text.upper()
        
        # Check for excluded terms
        for term in exclude_terms:
            if term in text_upper:
                return False
        
        # Check for street suffixes
        for suffix in street_suffixes:
            if text_upper.endswith(suffix):
                return False
        
        return True
    
    def _clean_name(self, name: str) -> str:
        """Normalize name for matching"""
        # Remove titles
        name = re.sub(r'\b(MR|MRS|MS|DR|JR|SR|III|II|IV)\.?\b', '', name, flags=re.IGNORECASE)
        
        # Remove special characters but keep spaces and hyphens
        name = re.sub(r'[^\w\s\-]', '', name)
        
        # Normalize whitespace
        name = ' '.join(name.split())
        
        # Convert to uppercase for matching
        return name.upper().strip()
    
    def parse_shipping_labels(self, pdf_bytes: bytes) -> List[Dict]:
        """Parse all shipping labels and extract names"""
        shipping_info = []
        
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    
                    # Try to extract name
                    name = self.extract_name_from_shipping_label(text)
                    
                    shipping_info.append({
                        'page': page_num,
                        'buyer_name': name,
                        'matched': False
                    })
                    
                    if not name:
                        print(f"⚠️ Could not extract name from shipping label page {page_num + 1}")
        
        except Exception as e:
            print(f"Error parsing shipping labels: {str(e)}")
            return []
        
        return shipping_info
    
    def match_labels_fuzzy(self, order_df: pd.DataFrame, shipping_info: List[Dict], 
                           threshold: int = 80) -> Tuple[List[Dict], List[Dict]]:
        """
        Match manufacturing labels with shipping labels using fuzzy matching
        Returns: (matched_pairs, unmatched_orders)
        """
        matched_pairs = []
        unmatched_orders = []
        
        # Get unique orders with buyer names
        orders = []
        for order_id in order_df['Order ID'].unique():
            order_data = order_df[order_df['Order ID'] == order_id].iloc[0]
            buyer_name = self._clean_name(order_data['Buyer Name'])
            orders.append({
                'order_id': order_id,
                'buyer_name': buyer_name,
                'item_count': len(order_df[order_df['Order ID'] == order_id])
            })
        
        # Match each order to a shipping label
        for order in orders:
            best_match = None
            best_score = 0
            best_index = -1
            
            # Try to match with available shipping labels
            for i, ship_info in enumerate(shipping_info):
                # Skip if already matched or no name extracted
                if ship_info['matched'] or not ship_info['buyer_name']:
                    continue
                
                # Calculate fuzzy match scores using multiple algorithms
                score1 = fuzz.ratio(order['buyer_name'], ship_info['buyer_name'])
                score2 = fuzz.token_sort_ratio(order['buyer_name'], ship_info['buyer_name'])
                score3 = fuzz.partial_ratio(order['buyer_name'], ship_info['buyer_name'])
                
                # Use the best score
                score = max(score1, score2, score3)
                
                if score > best_score:
                    best_score = score
                    best_match = ship_info
                    best_index = i
            
            # If we found a good match
            if best_match and best_score >= threshold:
                shipping_info[best_index]['matched'] = True
                matched_pairs.append({
                    'order_id': order['order_id'],
                    'buyer_name': order['buyer_name'],
                    'shipping_page': best_match['page'],
                    'shipping_name': best_match['buyer_name'],
                    'confidence': best_score,
                    'item_count': order['item_count']
                })
            else:
                unmatched_orders.append({
                    'order_id': order['order_id'],
                    'buyer_name': order['buyer_name'],
                    'item_count': order['item_count'],
                    'best_score': best_score
                })
        
        return matched_pairs, unmatched_orders
    
    def create_merged_pdfs(self, manufacturing_pdf: bytes, shipping_pdf: bytes,
                          matched_pairs: List[Dict], order_df: pd.DataFrame) -> Tuple[bytes, bytes]:
        """
        Create merged PDF for matched labels and separate PDF for unmatched
        Returns: (merged_pdf_bytes, unmatched_pdf_bytes)
        """
        # Read PDFs
        mfg_reader = PdfReader(io.BytesIO(manufacturing_pdf))
        ship_reader = PdfReader(io.BytesIO(shipping_pdf))
        
        # Create output PDFs
        merged_writer = PdfWriter()
        unmatched_writer = PdfWriter()
        
        # Build manufacturing page index by order ID
        mfg_page_index = {}
        current_page = 0
        for order_id in order_df['Order ID'].unique():
            item_count = len(order_df[order_df['Order ID'] == order_id])
            mfg_page_index[order_id] = list(range(current_page, current_page + item_count))
            current_page += item_count
        
        # Track which manufacturing pages have been merged
        merged_mfg_pages = set()
        
        # Add matched pairs to merged PDF
        for match in matched_pairs:
            order_id = match['order_id']
            
            # Add shipping label first
            merged_writer.add_page(ship_reader.pages[match['shipping_page']])
            
            # Add corresponding manufacturing labels
            if order_id in mfg_page_index:
                for mfg_page_num in mfg_page_index[order_id]:
                    merged_writer.add_page(mfg_reader.pages[mfg_page_num])
                    merged_mfg_pages.add(mfg_page_num)
        
        # Add unmatched manufacturing labels to separate PDF
        for page_num in range(len(mfg_reader.pages)):
            if page_num not in merged_mfg_pages:
                unmatched_writer.add_page(mfg_reader.pages[page_num])
        
        # Convert to bytes
        merged_output = io.BytesIO()
        merged_writer.write(merged_output)
        merged_output.seek(0)
        
        unmatched_output = io.BytesIO()
        if len(unmatched_writer.pages) > 0:
            unmatched_writer.write(unmatched_output)
            unmatched_output.seek(0)
            return merged_output.read(), unmatched_output.read()
        
        return merged_output.read(), b''

# Standalone function for easy integration
def merge_labels_with_improved_parser(manufacturing_pdf_bytes, shipping_pdf_bytes, order_dataframe, threshold=80):
    """
    Main function to merge labels using the improved parser
    
    Args:
        manufacturing_pdf_bytes: PDF bytes of manufacturing labels
        shipping_pdf_bytes: PDF bytes of shipping labels
        order_dataframe: DataFrame with columns 'Order ID' and 'Buyer Name'
        threshold: Minimum fuzzy match score (0-100)
    
    Returns:
        tuple: (merged_pdf_bytes, unmatched_pdf_bytes, statistics_dict)
    """
    merger = ImprovedLabelMerger()
    
    # Parse shipping labels
    print("📖 Parsing shipping labels...")
    shipping_info = merger.parse_shipping_labels(shipping_pdf_bytes)
    
    # Match labels
    print("🔗 Matching labels...")
    matched_pairs, unmatched_orders = merger.match_labels_fuzzy(
        order_dataframe, shipping_info, threshold
    )
    
    # Create merged PDFs
    print("📄 Creating merged PDFs...")
    merged_pdf, unmatched_pdf = merger.create_merged_pdfs(
        manufacturing_pdf_bytes, shipping_pdf_bytes, matched_pairs, order_dataframe
    )
    
    # Prepare statistics
    stats = {
        'total_orders': len(order_dataframe['Order ID'].unique()),
        'matched_orders': len(matched_pairs),
        'unmatched_orders': len(unmatched_orders),
        'match_rate': len(matched_pairs) / len(order_dataframe['Order ID'].unique()) * 100 if len(order_dataframe['Order ID'].unique()) > 0 else 0,
        'shipping_labels_processed': len(shipping_info),
        'matched_details': [(m['order_id'], m['buyer_name'], m['confidence']) for m in matched_pairs],
        'unmatched_details': [(u['order_id'], u['buyer_name']) for u in unmatched_orders]
    }
    
    # Print summary
    print("\n" + "=" * 60)
    print(f"✅ Matched: {stats['matched_orders']} orders ({stats['match_rate']:.1f}%)")
    print(f"⚠️ Unmatched: {stats['unmatched_orders']} orders")
    print(f"📊 Total Orders: {stats['total_orders']}")
    print(f"🏷️ Shipping Labels: {stats['shipping_labels_processed']}")
    
    if unmatched_orders:
        print("\n🔍 Unmatched Orders:")
        for order in unmatched_orders[:5]:  # Show first 5
            print(f"   - {order['order_id']}: {order['buyer_name']}")
        if len(unmatched_orders) > 5:
            print(f"   ... and {len(unmatched_orders) - 5} more")
    
    print("=" * 60)
    
    return merged_pdf, unmatched_pdf, stats

# Streamlit interface
def create_streamlit_app():
    """Streamlit UI for the improved label merger"""
    st.set_page_config(page_title="Label Merger", page_icon="🏷️", layout="wide")
    
    st.title("🏷️ Universal Label Merger - Improved Version")
    st.write("Successfully tested with UPS, USPS, and FedEx shipping labels")
    
    # File uploaders
    col1, col2, col3 = st.columns(3)
    with col1:
        mfg_file = st.file_uploader("📦 Manufacturing Labels PDF", type="pdf")
    with col2:
        ship_file = st.file_uploader("🚚 Shipping Labels PDF", type="pdf")
    with col3:
        order_file = st.file_uploader("📊 Order Data CSV", type="csv")
    
    # Settings
    with st.expander("⚙️ Settings"):
        threshold = st.slider(
            "Name Matching Threshold (%)", 
            min_value=60, 
            max_value=100, 
            value=80, 
            step=5,
            help="Lower = more matches but possibly incorrect, Higher = fewer matches but more accurate"
        )
        
        st.info("""
        **Supported Shipping Label Formats:**
        - ✅ UPS Ground / UPS Ground Saver
        - ✅ USPS Priority Mail / USPS Ground Advantage  
        - ✅ FedEx Express / FedEx Ground
        - ✅ Amazon shipping labels
        """)
    
    if mfg_file and ship_file and order_file:
        if st.button("🔄 Merge Labels", type="primary", use_container_width=True):
            with st.spinner("Processing labels..."):
                try:
                    # Read files
                    mfg_bytes = mfg_file.read()
                    ship_bytes = ship_file.read()
                    order_df = pd.read_csv(order_file)
                    
                    # Check required columns
                    if 'Order ID' not in order_df.columns or 'Buyer Name' not in order_df.columns:
                        st.error("❌ CSV must have 'Order ID' and 'Buyer Name' columns")
                        return
                    
                    # Merge labels
                    merged_pdf, unmatched_pdf, stats = merge_labels_with_improved_parser(
                        mfg_bytes, ship_bytes, order_df, threshold
                    )
                    
                    # Display results
                    st.success("✅ Processing Complete!")
                    
                    # Statistics
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Match Rate", f"{stats['match_rate']:.1f}%")
                    col2.metric("Matched Orders", stats['matched_orders'])
                    col3.metric("Unmatched Orders", stats['unmatched_orders'])
                    col4.metric("Total Orders", stats['total_orders'])
                    
                    # Download buttons
                    st.write("---")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label="⬇️ Download Merged Labels PDF",
                            data=merged_pdf,
                            file_name="merged_labels.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    
                    with col2:
                        if unmatched_pdf:
                            st.download_button(
                                label="⬇️ Download Unmatched Manufacturing Labels",
                                data=unmatched_pdf,
                                file_name="unmatched_manufacturing_labels.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                    
                    # Show match details
                    with st.expander("📋 View Match Details"):
                        if stats['matched_details']:
                            st.write("**Matched Orders:**")
                            for order_id, name, confidence in stats['matched_details'][:10]:
                                emoji = "🟢" if confidence >= 90 else "🟡"
                                st.write(f"{emoji} {order_id}: {name} ({confidence:.0f}% confidence)")
                            if len(stats['matched_details']) > 10:
                                st.write(f"... and {len(stats['matched_details']) - 10} more")
                        
                        if stats['unmatched_details']:
                            st.write("\n**Unmatched Orders:**")
                            for order_id, name in stats['unmatched_details'][:10]:
                                st.write(f"❌ {order_id}: {name}")
                            if len(stats['unmatched_details']) > 10:
                                st.write(f"... and {len(stats['unmatched_details']) - 10} more")
                    
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    st.exception(e)
    else:
        st.info("👆 Please upload all three files to begin")

if __name__ == "__main__":
    # If running with Streamlit
    create_streamlit_app()
