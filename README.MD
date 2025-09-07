# 📊 Adventure Works - Time Series Profitability Dashboard

This project is a comprehensive Excel-based dashboard that analyzes the profitability performance of the **Adventure Works** dataset over multiple years. It includes interactive insights into key metrics such as **revenue, COGS, gross profit, transaction volume**, and **profitability trends by time period**.

![Dashboard Preview](/Images/Dashboard.gif)

---

## 🧼 Data Cleaning Process

Data was cleaned and transformed using **Power Query Editor**. The key steps include:

- ✅ Changed column data types (e.g., to decimal, text)
- ✅ Extracted date elements (Month, Quarter, Weekday) from the `Date` table
- ✅ Calculated new columns in the Fact Table:
  - `Total Revenue = SalesAmount`
  - `COGS = TotalProductCost`
  - `Gross Profit = SalesAmount - TotalProductCost`
- ✅ Replaced null/NA values (e.g., product names, gender, territory names)
- ✅ Renamed and reordered columns
- ✅ Filtered and cleaned country/region values

![Power Query](/Images/Power_Query.PNG)

---

## 🔢 Excel Formulas Used

Several advanced Excel formulas were used for dynamic calculations, lookups, and highlighting:

### Example Formula:

### 🥇 Get Top 3 Gross Profit Values

```excel
=LARGE($T$27:$T$33,1)  ; 👉 Top 1
=LARGE($T$27:$T$33,2)  ; 👉 Top 2
=LARGE($T$27:$T$33,3)  ; 👉 Top 3
```

### 🎯 Highlight Top 3 Gross Profit Rows

```excel
=IFERROR(
  IFS(
    AB27=$W$27, AB27,
    AB27=$W$28, AB27,
    AB27=$W$29, AB27
  ),
"")
```

### 📊 Calculate % Contribution of Top 3

```excel
=SUM(AC27:AC33)/SUM(AB27:AB33)
```

### 📌 Other Formulas Used

- `XLOOKUP()` – For pulling values based on month names
- `INDEX + MATCH()` – For matching row-column intersections
- `IFS()` – For multi-conditional logic
- `IFERROR()` – To handle errors in calculated fields

### 📊 Pivot Tables Used

Several pivot tables were built to support the dashboard visuals and KPIs:

📌 Yearly Performance Overview: Revenue, Profit, COGS, Transactions
![Yearly_Performance](/Images/Pivot_Table_Yearly_Analysis.PNG)

📌 Monthly Profit Breakdown
![Monthly_Profit](/Images/Profit_by_Months.PNG)

📌 Profit by Day of Week

📌 Profit by Workdays vs. Weekend

📌 Quarterly Profit Trends

### 🔍 Dashboard KPIs & Insights

The dashboard provides real-time key performance indicators (KPIs) for evaluating profitability across different time dimensions. Here's a breakdown of the main metrics and insights:

![KPIs](/Images/KPis.PNG)

![KPIs_All](/Images/KPIs_PivotTables&Charts.PNG)

## 🖼️ Year-over-Year Comparison Example

Compare this year’s performance to last year using a clean and visual layout:
![Compare Last Year](/Images/Compare_Last_Year.PNG)

### 🧠 Entity Relationship Diagram

This project uses a star schema structure with dimensions and fact tables:
![Diagram](/Images/Diagram_View.PNG)

### 🔍 Key Highlights

🧾 Weekday vs Weekend:
![Weekday_vs_Weekend](/Images/Weedays_vs_Weekend.PNG)

📅 Quarterly Analysis:
![Quarter](/Images/Profit_by_Quarter.PNG)

### ✅ Features

- Dynamic Year/Month/Country filters

- Clear profit trendlines over time

- Easy comparison between current and previous years

- Visual KPIs for business decision-making
