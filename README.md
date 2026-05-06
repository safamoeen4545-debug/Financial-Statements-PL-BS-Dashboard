# Financial Analyst Project – Excel to Power BI Dashboard

## 📌 Project Overview

This project transforms raw accounting data (journals, ledger, trial balance) into an **interactive Power BI dashboard** that presents a complete financial picture: Profit & Loss, Balance Sheet, key ratios, and expense breakdowns.  
The source data is a fully functional Excel accounting file (`Copy of Financial analyst final.xlsx`) with charts of accounts, general entries, ledger, trial balance, income statement, and balance sheet.

**Goal:** Demonstrate how to move from Excel‑based financial reporting to a dynamic, visual, and single‑page dashboard using Power BI.

---

## 📁 Repository Contents

| File | Description |
|------|-------------|
| `Copy of Financial analyst final.xlsx` | Main Excel file with all accounting sheets (Charts of Accounts, General Entries, Ledger, Trial Balance, Income Statement, Balance Sheet) |
| `Financial_Dashboard.pbix` | Power BI dashboard file (to be created following the steps below) |
| `README.md` | This instruction file |

---

## 🗂️ Excel File Structure

The Excel file contains the following sheets:

| Sheet Name | Purpose |
|------------|---------|
| Charts of Accounts | Maps accounts to categories (Asset, Liability, Equity, Revenue, Expense) and to financial statement labels |
| General Entries | Journal entries with date, description, debit, credit, and narration |
| Ledger | Account‑wise summary of debits, credits, and running balance |
| Trial Balance | Final balances, sub‑account classification, and statement label |
| Income Statement | Pre‑built P&L using VLOOKUP formulas |
| Balance Sheet | Pre‑built balance sheet using VLOOKUP formulas |

> **Note:** The Excel file contains typos (e.g., "Balace Sheet" instead of "Balance Sheet") and VLOOKUP formulas that Power BI cannot evaluate. Cleaning steps are included below.

---

## 📊 Power BI Dashboard – What It Shows (One Page)


The final interactive dashboard includes:

| Visual | What It Displays |
|--------|------------------|
| **4 KPI Cards** | Total Revenue (1.23M), Net Profit (333K), Total Assets (6.23M), Total Liabilities & Equity (6.23M) |
| **Net Profit Margin Gauge** | 27.14% (target 20%) |
| **P&L by Account (Bar Chart)** | Sale, COGS, Salaries, Rent, Office Supplier, Entertainment Expense, Tax |
| **Revenue vs Expenses (Donut Chart)** | 57.85% Revenue, 42.15% Expenses |
| **Assets Breakdown (Column Chart)** | Cash, Bank, Inventory, Receivables (Mr. Rehman), Computer |
| **Liabilities & Equity (Column Chart)** | Mega Mart Vendor, Govt, Equity (shown positive) |
| **Expense Treemap** | COGS (largest), Tax, Salaries, Office Supplier, Rent, Entertainment Expense |
| **Waterfall Chart** | Bridge from Revenue to Net Profit (optional addition) |

---

## 🔧 Step‑by‑Step: From Excel to Power BI Dashboard

### 1. Prepare the Excel File

- Open `Copy of Financial analyst final.xlsx`
- Go to sheet **`Trail Balance`**
- **Delete the first 3 rows** (rows with "Ledger", "Trial Balance", and blank)
- **Copy all data** → **Paste Special → Values** (to remove VLOOKUP formulas)
- Correct the typo: change `"Balace Sheet"` to `"Balance Sheet"` in column `Label to FS` (or adjust measures later)
- Repeat for **`General Entries`** sheet (delete top rows, paste values) – optional for time trends
- Save as a new file (e.g., `Financial_Data_Cleaned.xlsx`)

### 2. Load Data into Power BI Desktop

- Open Power BI Desktop → **Get Data** → **Excel** → select the cleaned file
- In Navigator, select **`Trail Balance`** (and `General Entries` if needed)
- Click **Transform Data** (Power Query Editor)
- Set first row as headers (if needed)  
- Change data type of `Sum of Debit`, `Sum of Credit`, `Sum of Balance` to **Decimal Number**
- Click **Close & Apply**

### 3. Create Required DAX Measures

Go to **Modeling** → **New Measure** and add the following:

```DAX
Total Revenue = 
CALCULATE(
    SUM('Trail Balance'[Sum of Balance]),
    'Trail Balance'[Label to FS] = "Income Statement",
    'Trail Balance'[Sum of Balance] < 0
) * -1

Total Expenses = 
CALCULATE(
    SUM('Trail Balance'[Sum of Balance]),
    'Trail Balance'[Label to FS] = "Income Statement",
    'Trail Balance'[Sum of Balance] > 0
)

Net Profit = [Total Revenue] - [Total Expenses]

Net Profit Margin = DIVIDE([Net Profit], [Total Revenue], 0)

Total Assets = 
CALCULATE(
    SUM('Trail Balance'[Sum of Balance]),
    'Trail Balance'[Label to FS] = "Balance Sheet",
    'Trail Balance'[Sum of Balance] > 0
)

Total Liabilities & Equity = 
CALCULATE(
    SUM('Trail Balance'[Sum of Balance]),
    'Trail Balance'[Label to FS] = "Balance Sheet",
    'Trail Balance'[Sum of Balance] < 0
) * -1

Expense Amount = 
CALCULATE(
    SUM('Trail Balance'[Sum of Balance]),
    'Trail Balance'[Sub-Account] = "Expense"
)

Liability Balance = 
CALCULATE(
    SUM('Trail Balance'[Sum of Balance]),
    'Trail Balance'[Label to FS] = "Balance Sheet",
    'Trail Balance'[Sum of Balance] < 0
) * -1```
```
### 4. Build the Visuals (Arrange on One Page)
Use the following visual types and field assignments:

Visual	Fields / Settings
Card (4x)	Total Revenue, Net Profit, Total Assets, Total Liabilities & Equity
Gauge	Value = Net Profit Margin, Minimum = 0, Maximum = 0.5, Target = 0.2
Horizontal Bar Chart	Y‑axis = Account (filter Label to FS = "Income Statement"), X‑axis = Sum of Balance (or Expense Amount for expenses)
Donut Chart	Values = Total Revenue and Total Expenses
Clustered Column Chart (Assets)	X‑axis = Account (filter Sum of Balance > 0 and Label to FS = "Balance Sheet" and Account <> "Equity", "Govt", "Mega Mart-Vendor"), Y‑axis = Sum of Balance
Clustered Column Chart (Liabilities & Equity)	X‑axis = Account (filter Sum of Balance < 0 and Label to FS = "Balance Sheet"), Y‑axis = Liability Balance (measure)
Treemap	Group = Account (filter Sub-Account = "Expense"), Values = Expense Amount
Waterfall Chart (optional)	Category = Account (Income Statement in order), Measure = Sum of Balance (or Expense Amount for expenses with sign adjustment)
### 5. Apply Filters (Crucial for Correctness)
For P&L Bar Chart: Visual‑level filter Label to FS = "Income Statement"

For Expense Treemap: Visual‑level filter Sub-Account = "Expense"

For Assets Column Chart: Visual‑level filter Sum of Balance > 0 AND Label to FS = "Balance Sheet" AND Account NOT IN ("Equity", "Govt", "Mega Mart-Vendor")

For Liabilities & Equity Column Chart: Visual‑level filter Sum of Balance < 0 AND Label to FS = "Balance Sheet"

#### 6. Format for Professional Look
Numbers: Format as PKR 0,, M (e.g., 1.2M) for cards and axes

Net Profit Card: Conditional formatting – green if positive

Background: Light gray or white

Title: Add a text box at top: PLPL Financial Dashboard

##### 📈 Key Insights from the Dashboard
Metric	Value
Total Revenue	PKR 1,225,000
Net Profit	PKR 332,500
Net Profit Margin	27.14% ( >20% target)
Largest Expense	Cost of Goods Sold (PKR 490,000)
Largest Asset	Cash (PKR 4,425,000)
Total Assets	PKR 6,225,000
Total Liabilities + Equity	PKR 6,225,000 (balances)
#### 🛠️ How to Replicate This Project
Clone this repository

Download the Excel file (Copy of Financial analyst final.xlsx)

Follow the cleaning steps in Section 1

Open Power BI Desktop and load the cleaned data

Create measures and visuals as described

Arrange on one page and publish to Power BI Service (optional)

#### 📚 Learning Outcomes
By completing this project, you will learn:

How to clean Excel data for Power BI (remove formulas, fix headers, handle typos)

How to create financial measures (Revenue, Expenses, Net Profit, Margins, Assets, Liabilities)

How to design a single‑page interactive dashboard for financial reporting

How to troubleshoot common issues (blank cards, wrong filters, treemap not showing values)

How to tell a financial story using KPIs, bar charts, donuts, treemaps, and waterfalls

🤝 Connect with Me
Feel free to use this project in your portfolio. If you have questions or suggestions, open an issue or reach out.

Skills demonstrated: Power BI, DAX, Excel, Financial Accounting, Data Visualization, Dashboard Design

📄 License
This project is for educational and portfolio purposes. You may use and modify it freely.
