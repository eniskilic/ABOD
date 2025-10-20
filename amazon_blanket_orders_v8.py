"""
Airtable Base Setup Script - FIXED
Run this ONCE to create your tables and fields
"""

import requests
import json

# Your credentials
AIRTABLE_PAT = "pat3HPlu7bZzJep6t.2ea662c7b5e4f25f406969f987c5fdb9e15d5a2c6e933428934c8e5602ae7a68"
BASE_ID = "appxoNC3r5NSsTP3U"  # CORRECTED BASE ID

headers = {
    "Authorization": f"Bearer {AIRTABLE_PAT}",
    "Content-Type": "application/json"
}

print("🚀 Setting up your Airtable base...\n")

# Step 1: Create ORDERS table
print("📋 Creating ORDERS table...")
orders_table_payload = {
    "name": "Orders",
    "fields": [
        {"name": "Order ID", "type": "singleLineText"},
        {"name": "Order Date", "type": "singleLineText"},
        {"name": "Buyer Name", "type": "singleLineText"},
        {"name": "Status", "type": "singleSelect", "options": {
            "choices": [
                {"name": "New", "color": "blueBright"},
                {"name": "In Production", "color": "yellowBright"},
                {"name": "Quality Check", "color": "orangeBright"},
                {"name": "Packaging", "color": "purpleBright"},
                {"name": "Shipped", "color": "greenBright"},
                {"name": "Remake", "color": "redBright"}
            ]
        }},
        {"name": "Notes", "type": "multilineText"},
        {"name": "Shipping Address", "type": "multilineText"}
    ]
}

response = requests.post(
    f"https://api.airtable.com/v0/meta/bases/{BASE_ID}/tables",
    headers=headers,
    json=orders_table_payload
)

if response.status_code == 200:
    orders_table_id = response.json()["id"]
    print(f"✅ ORDERS table created! ID: {orders_table_id}\n")
else:
    print(f"❌ Error creating ORDERS table: {response.text}\n")
    print(f"Status code: {response.status_code}\n")
    exit()

# Step 2: Create ORDER LINE ITEMS table
print("📋 Creating ORDER LINE ITEMS table...")
line_items_payload = {
    "name": "Order Line Items",
    "fields": [
        {"name": "Buyer Name", "type": "singleLineText"},
        {"name": "Order ID", "type": "multipleRecordLinks", "options": {
            "linkedTableId": orders_table_id
        }},
        {"name": "Customization Name Placement", "type": "singleLineText"},
        {"name": "Quantity", "type": "number", "options": {"precision": 0}},
        {"name": "Blanket Color", "type": "singleLineText"},
        {"name": "Thread Color", "type": "singleLineText"},
        {"name": "Include Beanie", "type": "singleSelect", "options": {
            "choices": [
                {"name": "YES", "color": "greenBright"},
                {"name": "NO", "color": "grayBright"}
            ]
        }},
        {"name": "Gift Box", "type": "singleSelect", "options": {
            "choices": [
                {"name": "YES", "color": "greenBright"},
                {"name": "NO", "color": "grayBright"}
            ]
        }},
        {"name": "Gift Note", "type": "singleSelect", "options": {
            "choices": [
                {"name": "YES", "color": "greenBright"},
                {"name": "NO", "color": "grayBright"}
            ]
        }},
        {"name": "Gift Message", "type": "multilineText"},
        {"name": "Bobbin Color", "type": "singleSelect", "options": {
            "choices": [
                {"name": "Black Bobbin", "color": "grayDark1"},
                {"name": "White Bobbin", "color": "grayBright"}
            ]
        }},
        {"name": "Design Files", "type": "multipleAttachments"},
        {"name": "Status", "type": "singleSelect", "options": {
            "choices": [
                {"name": "Pending", "color": "grayBright"},
                {"name": "Cutting", "color": "blueBright"},
                {"name": "Embroidery", "color": "yellowBright"},
                {"name": "Quality Check", "color": "orangeBright"},
                {"name": "Complete", "color": "greenBright"}
            ]
        }},
        {"name": "Assigned To", "type": "singleLineText"},
        {"name": "Notes", "type": "multilineText"}
    ]
}

response = requests.post(
    f"https://api.airtable.com/v0/meta/bases/{BASE_ID}/tables",
    headers=headers,
    json=line_items_payload
)

if response.status_code == 200:
    line_items_table_id = response.json()["id"]
    print(f"✅ ORDER LINE ITEMS table created! ID: {line_items_table_id}\n")
else:
    print(f"❌ Error creating ORDER LINE ITEMS table: {response.text}\n")
    print(f"Status code: {response.status_code}\n")
    exit()

print("=" * 60)
print("✅ SUCCESS! Your Airtable base is ready!")
print("=" * 60)
print(f"\n📊 Tables created:")
print(f"   1. Orders (ID: {orders_table_id})")
print(f"   2. Order Line Items (ID: {line_items_table_id})")
print(f"\n🎯 Next steps:")
print(f"   1. Check your Airtable base to see the new tables")
print(f"   2. Use the updated Streamlit app to upload orders")
print(f"   3. For security: Regenerate your token after testing")
print("\n" + "=" * 60)
