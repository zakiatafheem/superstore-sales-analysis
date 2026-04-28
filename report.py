import pandas as pd
import matplotlib.pyplot as plt
import os

#  Create output folder
if not os.path.exists("output"):
    os.makedirs("output")

#  Load Dataset
df = pd.read_excel("Superstore_dataset.xlsx")

#  Convert Date
df['Order Date'] = pd.to_datetime(df['Order Date'])

#  Feature Engineering
df['Month'] = df['Order Date'].dt.to_period('M')

#  KPIs
total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
total_orders = df['Order ID'].nunique()

#  Analysis

# 1. Monthly Sales
monthly_sales = df.groupby('Month')['Sales'].sum().reset_index()

# 2. Category-wise Sales
category_sales = df.groupby('Category')['Sales'].sum().reset_index()

# 3. Sub-Category Sales
sub_category_sales = df.groupby('Sub-Category')['Sales'].sum().reset_index()

# 4. Profit Distribution
profit_stats = df['Profit'].describe()

# 5. Discount vs Profit
discount_profit = df[['Discount', 'Profit']]


#  VISUALIZATIONS

# Monthly Sales Line Plot
plt.figure()
plt.plot(monthly_sales['Month'].astype(str), monthly_sales['Sales'])
plt.xticks(rotation=45)
plt.title("Monthly Sales Trend")
plt.tight_layout()
plt.savefig("output/monthly_sales.png")
plt.close()

# Category Bar Plot
plt.figure()
plt.bar(category_sales['Category'], category_sales['Sales'])
plt.title("Category-wise Sales")
plt.tight_layout()
plt.savefig("output/category_sales.png")
plt.close()

# Sub-Category Bar Plot
plt.figure()
plt.barh(sub_category_sales['Sub-Category'], sub_category_sales['Sales'])
plt.title("Sub-Category Sales")
plt.tight_layout()
plt.savefig("output/subcategory_sales.png")
plt.close()

# Profit Boxplot
plt.figure()
plt.boxplot(df['Profit'])
plt.title("Profit Distribution")
plt.savefig("output/profit_boxplot.png")
plt.close()

# Discount vs Profit Scatter
plt.figure()
plt.scatter(df['Discount'], df['Profit'])
plt.xlabel("Discount")
plt.ylabel("Profit")
plt.title("Discount vs Profit")
plt.savefig("output/discount_vs_profit.png")
plt.close()

# INSIGHTS

top_category = category_sales.sort_values(by='Sales', ascending=False).iloc[0]
top_subcategory = sub_category_sales.sort_values(by='Sales', ascending=False).iloc[0]

insights = [
    f"Total Sales: {total_sales:.2f}",
    f"Total Profit: {total_profit:.2f}",
    f"Total Orders: {total_orders}",
    f"Top Category: {top_category['Category']}",
    f"Top Sub-Category: {top_subcategory['Sub-Category']}",
    "Higher discounts tend to reduce profit (see scatter plot)"
]

insights_df = pd.DataFrame(insights, columns=["Insights"])

# 📄 SAVE EXCEL REPORT

with pd.ExcelWriter("output/Superstore_Report.xlsx") as writer:
    monthly_sales.to_excel(writer, sheet_name="Monthly Sales", index=False)
    category_sales.to_excel(writer, sheet_name="Category Sales", index=False)
    sub_category_sales.to_excel(writer, sheet_name="SubCategory Sales", index=False)
    insights_df.to_excel(writer, sheet_name="Insights", index=False)

print("✅ Report + Charts Generated in 'output' folder!")