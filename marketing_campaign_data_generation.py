import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import datetime, timedelta

# Initialize Faker
fake = Faker()
Faker.seed(42)
np.random.seed(42)

# Configuration for Data Volume (Adjust these to hit 50MB)
NUM_CUSTOMERS = 5000
NUM_PRODUCTS = 100
NUM_TRANSACTIONS = 150000  # High volume for file size
NUM_CAMPAIGNS = 20
NUM_FEEDBACK = 5000

print("🚀 Starting PK House Data Factory...")

# --- 1. DIM_PRODUCTS ---
categories = {
    'Electronics': ['Smart Home', 'Wearables', 'Mobile Accessories'],
    'Fashion': ['Men', 'Women', 'Kids', 'Accessories'],
    'Home': ['Decor', 'Kitchen', 'Furniture']
}

products = []
for _ in range(NUM_PRODUCTS):
    cat = random.choice(list(categories.keys()))
    sub_cat = random.choice(categories[cat])
    base_cost = round(random.uniform(10, 500), 2)
    products.append({
        'SKU': f'PK-{fake.unique.bothify(text="??-####")}',
        'Product_Name': f"{fake.word().title()} {sub_cat} Item",
        'Category': cat,
        'Sub_Category': sub_cat,
        'Unit_Cost': base_cost,
        'Unit_Price': round(base_cost * random.uniform(1.3, 2.5), 2), # Healthy margin
        'Launch_Date': fake.date_between(start_date='-2y', end_date='today')
    })
df_products = pd.DataFrame(products)

# --- 2. DIM_CUSTOMERS ---
customers = []
for _ in range(NUM_CUSTOMERS):
    join_date = fake.date_between(start_date='-3y', end_date='today')
    customers.append({
        'Cust_ID': fake.unique.uuid4()[:8],
        'Name': fake.name(),
        'Gender': random.choice(['M', 'F', 'NB']),
        'Age': random.randint(18, 70),
        'City': random.choice(['Dubai', 'London', 'New York', 'Mumbai', 'Singapore', 'Berlin', 'Riyadh']),
        'Marital_Status': random.choice(['Single', 'Married', 'Divorced']),
        'Income_Bracket': random.choice(['Low', 'Medium', 'High', 'Very High']),
        'Join_Date': join_date,
        'Loyalty_Tier': random.choice(['Bronze', 'Silver', 'Gold', 'Platinum'])
    })
df_customers = pd.DataFrame(customers)

# --- 3. DIM_CAMPAIGNS ---
campaign_data = []
for i in range(NUM_CAMPAIGNS):
    start = fake.date_between(start_date='-1y', end_date='today')
    campaign_data.append({
        'Campaign_ID': f'CMP-{i+1:03d}',
        'Campaign_Name': f"PK {random.choice(['Summer', 'Winter', 'Eid', 'Black Friday', 'Flash'])} {random.choice(['Sale', 'Bash', 'Fest', 'Promo'])}",
        'Platform': random.choice(['Instagram', 'Google Ads', 'TikTok', 'Email', 'Influencer']),
        'Budget': round(random.uniform(5000, 50000), 2),
        'Start_Date': start,
        'End_Date': start + timedelta(days=random.randint(5, 30))
    })
df_campaigns = pd.DataFrame(campaign_data)

# --- 4. FACT_SALES (The Big Table) ---
print("Generating Transactions (this may take a moment)...")
transactions = []
customer_ids = df_customers['Cust_ID'].tolist()
product_skus = df_products['SKU'].tolist()

for _ in range(NUM_TRANSACTIONS):
    date = fake.date_between(start_date='-2y', end_date='today')
    qty = random.randint(1, 5)
    sku = random.choice(product_skus)
    # Seasonal Logic: More sales in Nov/Dec
    if date.month in [11, 12]:
        qty += random.randint(0, 3)
    
    transactions.append({
        'Order_ID': fake.unique.bothify(text='ORD-########'),
        'Date': date,
        'Cust_ID': random.choice(customer_ids),
        'SKU': sku,
        'Quantity': qty,
        'Channel': random.choice(['App', 'Website', 'In-Store']),
        'Payment_Method': random.choice(['Credit Card', 'Apple Pay', 'COD', 'BNPL']),
        'Campaign_ID': random.choice([None] + df_campaigns['Campaign_ID'].tolist()) # Some organic, some campaign
    })
df_sales = pd.DataFrame(transactions)

# Merge Price to calculate Revenue (for user convenience, though usually done in BI)
df_sales = df_sales.merge(df_products[['SKU', 'Unit_Price']], on='SKU', how='left')
df_sales['Total_Revenue'] = df_sales['Quantity'] * df_sales['Unit_Price']
df_sales.drop(columns=['Unit_Price'], inplace=True) # Keep Fact table clean

# --- 5. FACT_FEEDBACK (For NLP/RAG) ---
print("Generating Unstructured Text Data...")
feedback = []
sentiments = ['Positive', 'Neutral', 'Negative']
templates = {
    'Positive': ["Love the quality!", "Great service from PK House.", "Delivery was super fast.", "Best purchase ever."],
    'Neutral': ["It's okay.", "Good but expensive.", "Delivery was average.", "Product is fine."],
    'Negative': ["Terrible quality.", "Arrived damaged.", "Customer support was rude.", "Not worth the money."]
}

for _ in range(NUM_FEEDBACK):
    sent = random.choice(sentiments)
    feedback.append({
        'Review_ID': fake.unique.uuid4()[:8],
        'Cust_ID': random.choice(customer_ids),
        'SKU': random.choice(product_skus),
        'Rating': random.randint(4, 5) if sent == 'Positive' else (3 if sent == 'Neutral' else random.randint(1, 2)),
        'Review_Text': f"{random.choice(templates[sent])} {fake.sentence()}",
        'Date': fake.date_between(start_date='-1y', end_date='today')
    })
df_feedback = pd.DataFrame(feedback)

# --- EXPORT ---
print("💾 Saving to Excel (PK_House_Data.xlsx)...")
with pd.ExcelWriter('PK_House_Data.xlsx', engine='openpyxl') as writer:
    df_customers.to_excel(writer, sheet_name='DIM_CUSTOMERS', index=False)
    df_products.to_excel(writer, sheet_name='DIM_PRODUCTS', index=False)
    df_campaigns.to_excel(writer, sheet_name='DIM_CAMPAIGNS', index=False)
    df_sales.to_excel(writer, sheet_name='FACT_SALES', index=False)
    df_feedback.to_excel(writer, sheet_name='FACT_FEEDBACK', index=False)

print("✅ Done! File 'PK_House_Data.xlsx' created.")