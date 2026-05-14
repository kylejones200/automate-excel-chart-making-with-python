---
author: "Kyle Jones"
date_published: "August 23, 2025"
date_exported_from_medium: "November 10, 2025"
canonical_link: "https://medium.com/@kyle-t-jones/automate-excel-chart-making-with-python-d33f4303f95a"
---

# Automate Excel Chart Making with Python Excel remains the world's most common business tool. Analysts use it for
reports, charts, and dashboards every day. But if you've ever...

### Automate Excel Chart Making with Python
Excel remains the world's most common business tool. Analysts use it for reports, charts, and dashboards every day. But if you've ever updated the same report every week, copied the same formulas into multiple files, or redrawn the same chart for different regions, you know how painful and repetitive Excel can become.

Python changes that. With libraries like Pandas and Openpyxl, you can automate the entire workflow: filter data, build pivot tables, format worksheets, and even generate charts inside Excel automatically.

This tutorial shows how to create quarterly sales charts in Excel with Python. We'll start with one region, then scale up to generate polished reports for every region in the dataset.

### Why Automate Excel with Python?
Manually updating Excel reports is error-prone. It's easy to miscopy a formula, miss a row, or overwrite the wrong sheet. With Python, you can:

- Save time --- a script runs in seconds.
- Ensure consistency --- the same transformations every time.
- Scale easily --- produce reports for 1, 10, or 100 regions without more work.
- Extend Excel --- combine Python's analytical power with Excel's familiar interface.

Automation doesn't replace Excel --- it makes Excel better.

### Step 1: Load the Libraries
We'll use Pandas for data wrangling and Openpyxl for Excel charting.

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
```

### Step 2: Load the Data
For this example, we'll use a sample sales dataset hosted on GitHub.

``` 
df = pd.read_excel(
    'https://github.com/datagy/pivot_table_pandas/raw/master/sample_pivot.xlsx',
    parse_dates=['Date']
)
df.head()
```

### Step 3: Build a Pivot Table
Let's filter to the East region and build quarterly sales totals by product type.

``` 
filtered = df[df['Region'] == 'East']
quarterly_sales = pd.pivot_table(
    filtered,
    index=filtered['Date'].dt.quarter,
    columns='Type',
    values='Sales',
    aggfunc='sum'
)

print("Quarterly Sales Pivot Table:")
print(quarterly_sales.head())
```

### Step 4: Save the Pivot Table to Excel
``` 
file_path = '/Users/jnesnky/Downloads/test.xlsx'
quarterly_sales.to_excel(file_path, sheet_name='Quarterly Sales', startrow=3)
```

### Step 5: Format the Sheet and Add a Chart
We can now load the workbook, style it, and insert a bar chart directly into Excel.

``` 
wb = load_workbook(file_path)
sheet1 = wb['Quarterly Sales']


sheet1['A1'] = 'Quarterly Sales'
sheet1['A2'] = 'datagy.io'
sheet1['A4'] = 'Quarter'
sheet1['A1'].style = 'Title'
sheet1['A2'].style = 'Headline 2'
for i in range(5, 9):
    sheet1[f'B{i}'].style='Currency'
    sheet1[f'C{i}'].style='Currency'
    sheet1[f'D{i}'].style='Currency'
bar_chart = BarChart()
data = Reference(sheet1, min_col=2, max_col=4, min_row=4, max_row=8)
categories = Reference(sheet1, min_col=1, max_col=1, min_row=5, max_row=8)
bar_chart.add_data(data, titles_from_data=True)
bar_chart.set_categories(categories)
sheet1.add_chart(bar_chart, "F4")
bar_chart.title = 'Sales by Type'
bar_chart.style = 3
wb.save(filename=file_path)
```

### Step 6: Automate for All Regions
The real power comes when you scale. Instead of repeating the process by hand, you can loop over every region in the dataset and generate separate Excel reports with charts included.

``` 
regions = list(df['Region'].unique())
folder_path = '/Users/jnesnky/Downloads/test'


for region in regions:
    filtered = df[df['Region'] == region]
    quarterly_sales = pd.pivot_table(
        filtered,
        index=filtered['Date'].dt.quarter,
        columns='Type',
        values='Sales',
        aggfunc='sum'
    )
    file_path = f"{folder_path}{region}.xlsx"
    quarterly_sales.to_excel(file_path, sheet_name='Quarterly Sales', startrow=3)
    
    wb = load_workbook(file_path)
    sheet1 = wb['Quarterly Sales']
    
    sheet1['A1'] = 'Quarterly Sales'
    sheet1['A2'] = 'datagy.io'
    sheet1['A4'] = 'Quarter'
    sheet1['A1'].style = 'Title'
    sheet1['A2'].style = 'Headline 2'
    for i in range(5, 10):
        sheet1[f'B{i}'].style='Currency'
        sheet1[f'C{i}'].style='Currency'
        sheet1[f'D{i}'].style='Currency'
    bar_chart = BarChart()
    data = Reference(sheet1, min_col=2, max_col=4, min_row=4, max_row=8)
    categories = Reference(sheet1, min_col=1, max_col=1, min_row=5, max_row=8)
    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(categories)
    sheet1.add_chart(bar_chart, "F4")
    bar_chart.title = 'Sales by Type'
    bar_chart.style = 3
    wb.save(file_path)
```

### Visualizing the Results
Here's what one of the automatically generated charts looks like:


And here's a diagram showing how Python automates the workflow --- taking multiple datasets and producing polished Excel reports:


### Why This Matters
Think about the hours saved: instead of manually building reports for each region, you now have a script that can generate them all in one go. Beyond sales data, this approach works for finance reports, HR dashboards, inventory tracking, or any recurring business report.

Python doesn't replace Excel --- it supercharges it.

### Summary
In this tutorial, you learned how to:

- Load and filter data with Pandas
- Build pivot tables for analysis
- Format Excel sheets with Openpyxl
- Add charts directly into Excel
- Scale the process to every region in the dataset

The next time you face repetitive Excel work, don't copy and paste. Write a script once, and let Python do the work forever.
