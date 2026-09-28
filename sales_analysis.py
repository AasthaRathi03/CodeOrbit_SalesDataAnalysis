import pandas as pd

# ==========================================
# 1. LOAD DATASET
# ==========================================

df = pd.read_csv("data/sales_data.csv.csv")

print("===== FIRST 5 ROWS =====")
print(df.head())

# ==========================================
# 2. DATASET INFORMATION
# ==========================================

print("\n===== DATASET INFORMATION =====")
print(df.info())

print("\n===== DATASET SHAPE =====")
print(df.shape)

# ==========================================
# 3. DATA QUALITY CHECK
# ==========================================

print("\n===== MISSING VALUES =====")
print(df.isnull().sum())

print("\n===== DUPLICATE ROWS =====")
print(df.duplicated().sum())

# ==========================================
# 4. DATE CONVERSION
# ==========================================

df["Order Date"] = pd.to_datetime(df["Order Date"])

print("\n===== ORDER DATE DATA TYPE =====")
print(df["Order Date"].dtype)

# ==========================================
# 5. KEY SALES METRICS
# ==========================================

total_sales = df["Sales"].sum()
total_orders = df["Order ID"].nunique()
average_order_value = total_sales / total_orders
total_quantity = df["Quantity"].sum()
total_profit = df["Profit"].sum()

print("\n===== KEY SALES METRICS =====")
print(f"Total Sales: ${total_sales:,.2f}")
print(f"Total Orders: {total_orders:,}")
print(f"Average Order Value: ${average_order_value:,.2f}")
print(f"Total Quantity Sold: {total_quantity:,}")
print(f"Total Profit: ${total_profit:,.2f}")

# ==========================================
# 6. TOP 10 PRODUCTS
# ==========================================

top_products = (
    df.groupby("Product Name")
    .agg(
        Total_Sales=("Sales", "sum"),
        Quantity_Sold=("Quantity", "sum"),
        Total_Profit=("Profit", "sum"),
    )
    .sort_values("Total_Sales", ascending=False)
    .head(10)
)

print("\n===== TOP 10 PRODUCTS BY SALES =====")
print(top_products)

# ==========================================
# 7. CATEGORY-WISE ANALYSIS
# ==========================================

category_analysis = (
    df.groupby("Category")
    .agg(
        Total_Sales=("Sales", "sum"),
        Quantity_Sold=("Quantity", "sum"),
        Total_Profit=("Profit", "sum"),
        Number_of_Orders=("Order ID", "nunique"),
    )
    .sort_values("Total_Sales", ascending=False)
)

print("\n===== CATEGORY-WISE ANALYSIS =====")
print(category_analysis)

# ==========================================
# 8. SUB-CATEGORY ANALYSIS
# ==========================================

subcategory_analysis = (
    df.groupby("Sub-Category")
    .agg(
        Total_Sales=("Sales", "sum"),
        Quantity_Sold=("Quantity", "sum"),
        Total_Profit=("Profit", "sum"),
    )
    .sort_values("Total_Sales", ascending=False)
)

print("\n===== TOP SUB-CATEGORIES =====")
print(subcategory_analysis.head(10))

# ==========================================
# 9. REGION-WISE ANALYSIS
# ==========================================

region_analysis = (
    df.groupby("Region")
    .agg(
        Total_Sales=("Sales", "sum"),
        Quantity_Sold=("Quantity", "sum"),
        Total_Profit=("Profit", "sum"),
        Number_of_Orders=("Order ID", "nunique"),
    )
    .sort_values("Total_Sales", ascending=False)
)

print("\n===== REGION-WISE ANALYSIS =====")
print(region_analysis)

# ==========================================
# 10. MONTHLY SALES ANALYSIS
# ==========================================

df["Year-Month"] = df["Order Date"].dt.to_period("M")

monthly_sales = df.groupby("Year-Month").agg(
    Total_Sales=("Sales", "sum"),
    Total_Profit=("Profit", "sum"),
    Quantity_Sold=("Quantity", "sum"),
)

print("\n===== MONTHLY SALES ANALYSIS =====")
print(monthly_sales)

# ==========================================
# 11. YEAR-WISE ANALYSIS
# ==========================================

df["Year"] = df["Order Date"].dt.year

yearly_sales = df.groupby("Year").agg(
    Total_Sales=("Sales", "sum"),
    Total_Profit=("Profit", "sum"),
    Quantity_Sold=("Quantity", "sum"),
)

print("\n===== YEAR-WISE SALES ANALYSIS =====")
print(yearly_sales)

# ==========================================
# 12. SAVE ANALYSIS RESULTS
# ==========================================

with pd.ExcelWriter("sales_analysis_results.xlsx", engine="openpyxl") as writer:

    top_products.to_excel(writer, sheet_name="Top Products")
    category_analysis.to_excel(writer, sheet_name="Category Analysis")
    subcategory_analysis.to_excel(writer, sheet_name="Subcategory Analysis")
    region_analysis.to_excel(writer, sheet_name="Region Analysis")
    monthly_sales.to_excel(writer, sheet_name="Monthly Sales")
    yearly_sales.to_excel(writer, sheet_name="Yearly Sales")

# ==========================================
# 13. IMPORTANT FINDINGS
# ==========================================

top_product = top_products.index[0]
top_product_sales = top_products.iloc[0]["Total_Sales"]

top_category = category_analysis.index[0]
top_category_sales = category_analysis.iloc[0]["Total_Sales"]

top_region = region_analysis.index[0]
top_region_sales = region_analysis.iloc[0]["Total_Sales"]

best_month = monthly_sales["Total_Sales"].idxmax()
best_month_sales = monthly_sales.loc[best_month, "Total_Sales"]

best_year = yearly_sales["Total_Sales"].idxmax()
best_year_sales = yearly_sales.loc[best_year, "Total_Sales"]

print("\n===== KEY FINDINGS =====")
print(
    f"1. The product with the highest total sales was "
    f"'{top_product}' with sales of ${top_product_sales:,.2f}."
)

print(
    f"2. The highest-selling category was "
    f"'{top_category}' with sales of ${top_category_sales:,.2f}."
)

print(
    f"3. The region with the highest total sales was "
    f"'{top_region}' with sales of ${top_region_sales:,.2f}."
)

print(
    f"4. The month with the highest sales was "
    f"{best_month} with sales of ${best_month_sales:,.2f}."
)

print(
    f"5. The year with the highest sales was "
    f"{best_year} with sales of ${best_year_sales:,.2f}."
)

print("\nAnalysis completed successfully.")
print("Results saved to sales_analysis_results.xlsx")
