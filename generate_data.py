import pandas as pd
from datetime import datetime, timedelta
import numpy as np

# --- Configuration ---
NUM_SALES_RECORDS = 5000
NUM_CUSTOMER_RECORDS = 1000
PRODUCTS = {
    "Laptop": 1200, "Mouse": 25, "Keyboard": 75, "Monitor": 300,
    "USB Hub": 15, "Webcam": 50, "Docking Station": 150
}
REGIONS = ["North", "South", "East", "West"]

# --- Helper Functions ---
def get_random_date(start_date, end_date):
    """Generate a random date between two dates."""
    return start_date + timedelta(
        days=np.random.randint(0, (end_date - start_date).days)
    )

# --- 1. Generate Sales Data Sheet ---
print("Generating Sales Data...")
sales_data = []
start_date = datetime(2023, 1, 1)
end_date = datetime(2024, 12, 31)

for i in range(1, NUM_SALES_RECORDS + 1):
    product_name = np.random.choice(list(PRODUCTS.keys()))
    quantity = np.random.randint(1, 5)
    unit_price = PRODUCTS[product_name]
    
    # Intentionally create variation in column names for a later challenge
    revenue = quantity * unit_price
    if i % 10 == 0:
        # Every 10th record uses a different date format (string)
        order_date = get_random_date(start_date, end_date).strftime("%m-%d-%Y")
    else:
        order_date = get_random_date(start_date, end_date)

    sales_data.append({
        "OrderID": 1000 + i,
        "Order Date": order_date,
        "Product": product_name,
        "Quantity": quantity,
        "Unit Price": unit_price,
        # Inconsistent column name: 'Revenue' vs 'sale_amount'
        "Revenue" if i % 2 == 0 else "sale_amount": revenue,
        "Region": np.random.choice(REGIONS),
        "CustomerID": np.random.randint(1, NUM_CUSTOMER_RECORDS + 1)
    })

sales_df = pd.DataFrame(sales_data)
# Convert 'Order Date' to datetime to ensure consistent type
sales_df['Order Date'] = pd.to_datetime(sales_df['Order Date'], errors='coerce')
# Fill NaN for the revenue columns created by the inconsistent naming
sales_df['Revenue'] = sales_df['Revenue'].fillna(sales_df['sale_amount'])
sales_df.drop(columns=['sale_amount'], inplace=True)


# --- 2. Generate Customer Data Sheet ---
print("Generating Customer Data...")
customer_data = []
for i in range(1, NUM_CUSTOMER_RECORDS + 1):
    customer_data.append({
        "customer_id": i, # snake_case for variety
        "FirstName": f"Customer{i}", # camelCase for variety
        "LastName": "User",
        "Join Date": get_random_date(datetime(2020, 1, 1), end_date),
        "Last Order Date": None # We can fill this later
    })
customers_df = pd.DataFrame(customer_data)

# Simulate Last Order Date
last_orders = sales_df.groupby('CustomerID')['Order Date'].max().reset_index()
last_orders.rename(columns={'CustomerID': 'customer_id', 'Order Date': 'Last Order Date'}, inplace=True)

# Convert string dates in last_orders to datetime
last_orders['Last Order Date'] = pd.to_datetime(last_orders['Last Order Date'], errors='coerce')

customers_df.set_index('customer_id', inplace=True)
last_orders.set_index('customer_id', inplace=True)
customers_df.update(last_orders)
customers_df.reset_index(inplace=True)


# --- 3. Generate Product Info Sheet ---
print("Generating Product Info...")
products_df = pd.DataFrame(list(PRODUCTS.items()), columns=["Product Name", "Price"])


# --- 4. Create an Empty Sheet (for edge case handling) ---
print("Creating an empty sheet...")
empty_df = pd.DataFrame()


# --- Write to Excel File ---
file_path = "sales_data.xlsx"
print(f"Writing all data to {file_path}...")
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    sales_df.to_excel(writer, sheet_name='Sales', index=False)
    customers_df.to_excel(writer, sheet_name='Customers', index=False)
    products_df.to_excel(writer, sheet_name='Product Info', index=False)
    empty_df.to_excel(writer, sheet_name='Empty Sheet', index=False)

print("Done! Sample Excel file 'sales_data.xlsx' has been created.")
