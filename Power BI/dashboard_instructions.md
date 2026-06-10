# Power BI Dashboard Instructions

This guide provides step-by-step instructions for building a comprehensive 4-page Power BI dashboard from scratch using the provided Star Schema dataset (`powerbi_dataset.xlsx`).

---

## 1. How to Import the Data
1. Open **Power BI Desktop**.
2. Click **Get Data** > **Excel workbook**.
3. Browse to and select `data/powerbi_dataset.xlsx`.
4. In the Navigator window, check the boxes next to all 5 sheets:
   - `Fact_Transactions`
   - `Dim_Product`
   - `Dim_Date`
   - `Dim_City`
   - `KPI_Summary`
5. Click **Load**. (Wait for the data to finish loading into the model).

---

## 2. Model View Relationships
Navigate to the **Model view** (the third icon on the left sidebar). Power BI may have auto-detected some relationships, but you must ensure the following connections exist exactly as specified:

| From (Many side `*`) | To (One side `1`) | Cross Filter Direction |
| --- | --- | --- |
| `Fact_Transactions[product_id]` | `Dim_Product[product_id]` | Single |
| `Fact_Transactions[Date]` | `Dim_Date[Date]` | Single |
| `Fact_Transactions[city]` | `Dim_City[City]` | Single |

*Note: Ensure that the relationship lines show an asterisk `*` on the `Fact_Transactions` side and a `1` on the Dimension tables side. Delete any auto-generated relationships that conflict with these.*

---

## 3. DAX Measures
Navigate to the **Data view** or **Report view**. Right-click on the `Fact_Transactions` table and select **New Measure** for each of the following (copy and paste the exact DAX code):

**Total Revenue**
```dax
Total Revenue = SUM(Fact_Transactions[revenue])
```

**Total Orders**
```dax
Total Orders = COUNT(Fact_Transactions[transaction_id])
```

**Avg Order Value**
```dax
Avg Order Value = DIVIDE([Total Revenue], [Total Orders])
```

**Total Customers**
```dax
Total Customers = DISTINCTCOUNT(Fact_Transactions[customer_id])
```

**Revenue Online**
```dax
Revenue Online = CALCULATE([Total Revenue], Fact_Transactions[Channel] = "online")
```

**Revenue Offline**
```dax
Revenue Offline = CALCULATE([Total Revenue], Fact_Transactions[Channel] = "offline")
```

**Online Revenue %**
```dax
Online Revenue % = DIVIDE([Revenue Online], [Total Revenue], 0) * 100
```

**MoM Revenue Growth**
```dax
MoM Revenue Growth = 
VAR CurrentMonth = [Total Revenue]
VAR PrevMonth = CALCULATE([Total Revenue], DATEADD(Dim_Date[Date], -1, MONTH))
RETURN DIVIDE(CurrentMonth - PrevMonth, PrevMonth, 0) * 100
```

**Top City**
```dax
Top City = TOPN(1, VALUES(Fact_Transactions[city]), [Total Revenue])
```

**YTD Revenue**
```dax
YTD Revenue = TOTALYTD([Total Revenue], Dim_Date[Date])
```

---

## 4. Page 1 — Revenue Overview
Create a new page named **Revenue Overview** and add the following visuals:

- **KPI Card:** Add `Total Revenue`
- **KPI Card:** Add `Total Orders`
- **KPI Card:** Add `Avg Order Value`
- **KPI Card:** Add `Total Customers`
- **Line Chart (Revenue by Month):**
  - **X-axis:** `Dim_Date[Month_Name]` *(Pro tip: Select the Month_Name column in the data pane, and click "Sort by column" > "Month" so it sorts chronologically instead of alphabetically)*
  - **Y-axis:** `Total Revenue`
- **Donut Chart (Revenue by Channel):**
  - **Legend:** `Fact_Transactions[Channel]`
  - **Values:** `Total Revenue`
- **Slicers:**
  - `Dim_Date[Year]`
  - `Fact_Transactions[Category]`

---

## 5. Page 2 — Product Performance
Create a new page named **Product Performance** and add the following visuals:

- **Bar Chart (Horizontal):** Top 10 Products by Revenue
  - **Y-axis:** `Fact_Transactions[product_name]`
  - **X-axis:** `Total Revenue`
  - *(Apply a Top N filter in the filters pane: Show Top 10 by Total Revenue)*
- **Table (Product Details):**
  - **Columns:** `product_name`, `Total Revenue`, `Total Orders`, Average of `discount` (Ensure it is set to Average, not Sum)
- **Treemap (Revenue by Category):**
  - **Category:** `Fact_Transactions[Category]`
  - **Values:** `Total Revenue`
- **Slicers:**
  - `Fact_Transactions[city]`
  - `Fact_Transactions[Channel]`

---

## 6. Page 3 — Regional Insights
Create a new page named **Regional Insights** and add the following visuals:

- **Bar Chart (Revenue by City):**
  - **X-axis:** `Fact_Transactions[city]`
  - **Y-axis:** `Total Revenue`
  - *(Sort visual by Total Revenue descending)*
- **Map Visual:**
  - **Location:** `Fact_Transactions[city]`
  - **Bubble Size:** `Total Revenue`
- **Table (City Summary):**
  - **Columns:** `city`, `Total Revenue`, `Total Orders`, `Top Product` (Using a custom top product measure or dragging product_name and filtering top 1)
- **Slicers:**
  - `Dim_Date[Year]`
  - `Fact_Transactions[Category]`

---

## 7. Page 4 — Category Trends
Create a new page named **Category Trends** and add the following visuals:

- **Clustered Column Chart (Revenue by Category per Month):**
  - **X-axis:** `Dim_Date[Month_Name]`
  - **Y-axis:** `Total Revenue`
  - **Legend:** `Fact_Transactions[Category]`
- **Line Chart (Monthly Trend per Category):**
  - **X-axis:** `Dim_Date[Month_Name]`
  - **Y-axis:** `Total Revenue`
  - **Legend:** `Fact_Transactions[Category]`
- **Matrix (Category vs City):**
  - **Rows:** `Fact_Transactions[Category]`
  - **Columns:** `Fact_Transactions[city]`
  - **Values:** `Total Revenue`
- **Slicers:**
  - `Dim_Date[Quarter]`
  - `Dim_Date[Year]`

---

## 8. Formatting Tips

- **Color Theme:** Apply a consistent, professional theme like **"Executive"** or **"Storm"** (found under the View ribbon > Themes). Avoid overly bright palettes; stick to cool blues, greys, and teals for a modern analytical look.
- **KPI Cards:** Increase the data label font size to **30pt+** and use a bold font (e.g., DIN or Segoe UI Semibold). Add a subtle drop shadow to the cards (Visual format > General > Effects > Shadow) to give depth to your dashboard.
- **Numbers:** Ensure all revenue DAX measures are formatted as **Currency (₹ or Rs., 0 decimal places)** in the measure tools ribbon at the top of the screen. Format percentages to 1 decimal place.
- **Conditional Formatting:** In the Matrix and Table visuals, right-click the `Total Revenue` field in the visual build pane and select **Conditional Formatting** > **Data Bars** or **Background Color Scales**. This makes high values instantly recognizable without the user having to read every number.
- **Titles:** Turn on titles for every chart. Center-align them, give them a distinct background color (e.g., dark grey/blue), and use white text to create visual boundaries.
