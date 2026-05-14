# 👥 PROJECT 1 (Employee Data Entry System)

**Automated Excel Form for Employee Onboarding & Payroll Management**

## 📋 Project Overview

This project is a **custom automated Excel data entry form** designed to streamline the onboarding process for newly joined employees. It efficiently captures employee personal details, department information, and payroll components (Basic Salary, HRA, Medical, etc.), with automatic submission to a structured database.

The tool enhances data accuracy, reduces manual work, and maintains a professional HR workflow.

## 📸 Screenshots

### 1. Main Data Entry Form
![Data Entry Form](https://github.com/wisemansg/advancedexcel/raw/main/assets/Data%20Entry%20Form%201.png)

### 2. Form Interface View
![Form View 2](https://github.com/wisemansg/advancedexcel/raw/main/assets/Data%20Entry%20Form%202.png)

### 3. Database Sheet (All Records)
![Database Sheet](https://github.com/wisemansg/advancedexcel/raw/main/assets/Database.png)

## 📥 Download

 [Download Automated Data Entry Form.xlsm](https://raw.githubusercontent.com/wisemansg/advancedexcel/main/assets/Automated%20Data%20Entry%20Form%20for%20newly%20joint%20employees%20in%20the%20company.xlsm)

## ✨ Key Features

- User-friendly and clean interface
- One-click **SUBMIT** button with VBA automation
- Automatic **Net Pay** calculation
- Location selection (Site / Office)
- Sheet protection (only input fields are editable)
- All records stored neatly in the **Database** sheet

## 🛠️ How It Works

1. Open the downloaded Excel file
2. Fill in the employee details on the **Form** sheet
3. Click the **SUBMIT** button
4. Data is automatically saved to the **Database** sheet
5. Form resets for the next entry

## 🔒 Security Features

- Saved as **.xlsm** (Macro-Enabled Workbook)
- Sheet protection enabled
- Only form input cells are unlocked
- Formulas and layout are locked

## 🧰 Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Form Controls & ActiveX
- Sheet Protection
- Excel Tables

## 📌 Future Enhancements

- Duplicate employee code validation
- Automatic email notification
- Search & filter in Database
- Summary Dashboard
- PDF export option

**Made with ❤️ for efficient HR processes**

---
# 👥 PROJECT 2 (Automated Excel Finance Tracker)

**Smart Personal Finance Management Dashboard in Excel**  
*Track spending, auto-classify transactions, monitor savings goals, and visualize your financial health — all automatically.*

## 📋 Project Overview

This **Automated Finance Tracker** eliminates manual data entry by allowing you to simply paste your bank transactions. Excel then automatically classifies every transaction, updates PivotTables, refreshes charts, and manages your savings goals with priority-based allocation. Built for clarity, automation, and ease of use.

## 📥 Download

- **Direct Download:**  
  👉 [Download Automated Finance Tracker.xlsm](https://raw.githubusercontent.com/wisemansg/advancedexcel/main/assets/automated_finance_tracker.xlsx)

## 📸 Project Screenshots

### Transaction Table (Data Entry Area)
![Transaction Table](https://github.com/wisemansg/advancedexcel/raw/main/assets/Transaction%20Table.png)

### Category Classification Table
![Category Classification](https://github.com/wisemansg/advancedexcel/raw/main/assets/Category%20Classification.png)

### PivotTables Analysis & Working Sheet
![PivotTables Analysis](https://github.com/wisemansg/advancedexcel/raw/main/assets/PivotTables%20Analysis%20and%20Working%20Sheet.png)

### Net Income PivotTable
![Net Income PivotTable](https://github.com/wisemansg/advancedexcel/raw/main/assets/Net%20Income%20PivotTable.png)

### Dashboard 1 – Overview
![Dashboard 1](https://github.com/wisemansg/advancedexcel/raw/main/assets/Dashboard%201.png)

### Dashboard 2 – Detailed View
![Dashboard 2](https://github.com/wisemansg/advancedexcel/raw/main/assets/Dashboard%202.png)

### Waterfall Chart – Where the Money Really Goes
![Waterfall Chart](https://github.com/wisemansg/advancedexcel/raw/main/assets/Waterfall%20Chart%20(Where%20the%20money%20really%20goes).png)

### Treemap Chart – Category Spending Comparison
![Treemap Chart](https://github.com/wisemansg/advancedexcel/raw/main/assets/Treemap%20Chart%20(The%20Fastest%20Way%20to%20Compare%20Category%20Spending).png)

### Savings Goals Tracker
![Savings Goals](https://github.com/wisemansg/advancedexcel/raw/main/assets/Savings%20Goals%20(Automatically%20Allocate%20Savings%20by%20Priority).png)

## ✨ Key Features

- One-click transaction import (paste from bank)
- **Automatic transaction classification** using XLOOKUP + SEARCH
- Dynamic PivotTables and Charts
- Waterfall & Treemap visualizations
- Automated Savings Goals tracker with priority allocation
- Clean Dashboard with real-time insights

## 🛠️ Core Formulas Used

1. Income/(Expense) Column

=[@[Credit (Income)]] - [@[Debit (Spend)]]

2. Automatic Classification (Subcategory, Category, Category Type)

=XLOOKUP(
    TRUE,
    ISNUMBER(SEARCH(Categories[Description],[@Description])),
    Categories[Subcategory],
    "Uncategorised"
)

3. Savings Goals – Total Net Savings

=-SUMIFS(tblTrans[Income/(Expense)], tblTrans[Subcategory], "Savings")

4. Cumulative Target

=SUMIFS([Target], [Priority], "<=" & [@Priority])

5. Allocated Savings (Using LET Function)

=LET(
    total,  TotalSavings,
    prev,   [@[Cum Target]] - [@Target],
    avail,  total - prev,
    MAX(0, MIN([@Target], avail))
)

6. % Funded
=IF([@Target]=0, 0, [@Allocated]/[@Target])

7. Months to Goal
=IF(
    TargetSavedMonthly<=0, "",
    MAX(0, ROUNDUP(([@[Cum Target]] - TotalSavings) / TargetSavedMonthly, 0))
)


## 🛠️ How It Works

1. **Paste transactions** into the `tblTrans` table on the Transactions sheet.
2. **Auto-classification** happens instantly using XLOOKUP and keyword matching.
3. **PivotTables** update automatically with the new data.
4. **Dashboards & Charts** refresh with one click using **Refresh All**.
5. **Savings Goals** automatically allocate available savings by priority order.

## 🧰 Technologies & Techniques Used

- Excel Tables (`tblTrans`, `tblCat`, `tblGoals`)
- Structured References
- XLOOKUP + SEARCH for smart classification
- LET Function for complex calculations
- Advanced PivotTables
- Dynamic Charts (Waterfall, Treemap, Donut)
- Sheet Protection & Named Ranges

## 📌 Future Enhancements

- Power Query auto-import from bank CSV
- Monthly budget comparison
- Expense forecasting
- Interactive slicers on dashboard

**Made with ❤️ to bring financial clarity and control**
---
# 📅 PROJECT 3 (Excel Task Tracker)

**Dynamic • Visual • Zero Macros**  
A powerful, fully automated **Assignment & Task Management System** in Excel — built for students and professionals who need clarity and control over their deadlines.

## 📥 Download

- **Direct Download:**  
  👉 [Download Excel Task Tracker.xlsx](https://raw.githubusercontent.com/wisemansg/advancedexcel/main/assets/excel_task_tracker.xlsx)

## 📸 Project Screenshots

### 1. Lists Sheet – Dynamic Drop-down Sources
![Lists for Drop-downs](https://github.com/wisemansg/advancedexcel/raw/main/assets/Lists%20(for%20drop-downs).png)

### 2. Main Task Tracker Table
![Task Tracker 1](https://github.com/wisemansg/advancedexcel/raw/main/assets/Task%20Tracker%201.png)

### 3. Tracker with Active Slicers & Conditional Formatting
![Task Tracker 2](https://github.com/wisemansg/advancedexcel/raw/main/assets/Task%20Tracker%202.png)

### 4. Tracker with Status Filters Applied
![Task Tracker 3](https://github.com/wisemansg/advancedexcel/raw/main/assets/Task%20Tracker%203.png)

### 5. Schedule View – Mini Gantt Calendar (Overview)
![Task Schedule 1](https://github.com/wisemansg/advancedexcel/raw/main/assets/Task%20Schedule%20(Mini%20Calendar)%201.png)

### 6. Schedule View – Due Date Highlights
![Task Schedule 2](https://github.com/wisemansg/advancedexcel/raw/main/assets/Task%20Schedule%20(Mini%20Calendar)%202.png)

## ✨ Key Features

- Smart drop-downs powered by dynamic lists
- Automatic **Days Available** countdown with overdue alerts
- Beautiful conditional formatting (color-coded status + strikethrough)
- Interactive **Slicers** for instant filtering
- Live **Mini Gantt Schedule** view using dynamic arrays
- Fully responsive — add tasks and everything updates automatically

## 🛠️ Core Formulas Used

1. Days Available (Countdown)
=[@Due] - TODAY()

2. Dynamic Filtered & Sorted Schedule (Active Tasks)
=SORT(
    FILTER(Tasks[[Task]:[Time Req'd]], Tasks[Status]<>"Completed"),
    4
)

3. Auto-Generated Date Row for Calendar
=SEQUENCE(, 
    MAX(Tasks[Due]) - MIN(Tasks[Due]) + 1, 
    MIN(Tasks[Due])
)

4. Day of Week Row (Spill Reference)
=G5#     // References the SEQUENCE row above


