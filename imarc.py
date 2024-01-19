from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd

# Step 1: Web Scraping using Selenium
url = "http://www.sunsirs.com/futures-price-2023-0927-daily.html"
driver = webdriver.Chrome()  # You may need to adjust the path to your Chrome driver

driver.get(url)

# Assuming the table has a class, replace 'your_table_class' with the actual class
table = driver.find_element(By.CLASS_NAME, "xnn-tablea")
rows = table.find_elements(By.TAG_NAME, 'tr')

data = []
for row in rows:
    cols = row.find_elements(By.TAG_NAME, 'td')
    cols = [col.text for col in cols]
    data.append(cols)

driver.quit()

# Extract 'end_date' values using index
end_date_values = [row[3] for row in data[1:]]  # Assuming 'end_date' is the fourth column, skipping header row


# Creating a Pandas DataFrame
columns = ["Commodity", "Sectors", "start_date", "end_date", 'Change']  # Replace with actual column names

df = pd.DataFrame(data[1:], columns=columns)  # Skip the header row

# Include the extracted 'end_date' values in the DataFrame
df['end_date'] = end_date_values

# Filter out rows where 'Commodity' is missing or consists only of whitespace
df = df[df['Commodity'].str.strip() != '']


# Convert 'end_date' values to numeric, handling non-numeric characters
df["end_date"] = pd.to_numeric(df["end_date"].str.replace('[^\d.]', '', regex=True), errors='coerce')


# Step 3: Data Analysis
# Count the total number of rows
total_rows = len(df)

# Find the commodity with the highest daily closing price
max_price_row = df.loc[df["end_date"].idxmax()]
max_commodity = max_price_row["Commodity"]
max_price = max_price_row["end_date"]

# Save the data to an Excel workbook
excel_file = "output_data.xlsx"
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    writer._save()

# Display results
print(f"\nTotal number of rows: {total_rows}")
print(f"Commodity with highest daily closing price: {max_commodity}, Price: {max_price}")
