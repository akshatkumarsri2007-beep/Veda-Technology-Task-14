**#Veda Technology Task 14**
# Retail Sales Summary in Excel

A basic Excel business summary that calculates **total sales, average sales, and transaction count** from retail sales data using formulas.

##  Objective

Build basic business summaries by turning raw transaction data into a few clear KPIs a stakeholder can read at a glance.

## Tools Used

- **Microsoft Excel**
  - `SUM`, `AVERAGE`, `COUNTA` formulas
  - Number / currency formatting

##  Dataset

- **Source:** Retail Sales sample dataset (prepared for practice)
- **Size:** 200 rows, 9 columns
- **Time range:** January – August 2026
- **Key fields:** Transaction ID, Date, Store, Category, Product, Quantity, Unit Price, Payment Mode, Sales Amount

##  Summary Sheet Features

## KPI Cards
- **Total Sales** — sum of all sales amounts → `=SUM('Retail Sales Data'!I2:I201)`
- **Average Sales** — average value of a single transaction → `=AVERAGE('Retail Sales Data'!I2:I201)`
- **Transaction Count** — number of transactions → `=COUNTA('Retail Sales Data'!A2:A201)`

## Results

| KPI | Value |
|---|---|
| Total Sales | ₹5,60,994 |
| Average Sales | ₹2,804.97 |
| Transaction Count | 200 |

## Verification
- Totals were verified manually using Excel's status bar (Sum, Average, Count)
- Cross-check: `Total Sales ÷ Transaction Count = Average Sales` → 5,60,994 ÷ 200 = 2,804.97

##  Deliverables

| File | Description |
|---|---|
| `Retail_Sales_Data.xlsx` | Dataset + Summary sheet with 3 KPIs |
| `Basic_Sales_Summary_Report.pdf` | 2-page report explaining the KPIs, breakdowns, and insights |

##  Key Insights

- Total sales of ₹5,60,994 were generated across 200 transactions
- Electronics and Home & Kitchen are the top categories, together making up about 68% of sales
- Bengaluru is the best-performing store, while Mumbai is the lowest
- UPI is the most used payment mode (37.5% of sales)
- Sales peak in May and are lowest in February
- Grocery contributes only ~5% of sales because of its low unit prices

##  How to Use

1. Download `Retail_Sales_Data.xlsx`
2. Open it in Microsoft Excel
3. Go to the **Summary** sheet to view the KPIs
4. Click any KPI cell to see the formula behind it
