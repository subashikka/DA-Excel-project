# Retail Sales Analysis Using Microsoft Excel

# Excel-Task

This repository contains a practical Microsoft Excel project focused on retail sales analysis. It demonstrates how Excel formulas can be used to summarize sales, identify high- and low-value transactions, classify sales records, compare regional performance, and extract information from transaction codes.

## 📊 Project Overview

The objective of this project is to analyze retail sales transactions using Microsoft Excel. The project applies basic statistical, logical, conditional, and text functions to transform raw sales data into meaningful business information.

The analysis focuses on sales performance, transaction values, regional sales totals, and transaction-code details.

**Dataset size:** 34 records with 6 original columns, following the structure of the original project. The retail sales scenario and column values are illustrative and should be adapted to the actual dataset.

## 🎯 Project Objectives

- Calculate total sales, transaction count, and average sales value.
- Identify the minimum and maximum transaction amounts.
- Classify transactions based on sales value using the IF function.
- Calculate region-wise sales totals using SUMIF.
- Count transactions below a specified sales threshold using COUNTIF.
- Extract day, month, and branch code from transaction IDs using LEFT, MID, and RIGHT.
- Develop practical Excel skills for retail sales reporting and analysis.

## 🛠️ Tools Used

- Microsoft Excel
- Excel formulas and functions
- Retail sales dataset

## 📂 Dataset Description

The dataset contains the following columns:

| Column | Field Name    | Description                                |
| ------ | ------------- | ------------------------------------------ |
| A      | Order Code    | Transaction identifier                     |
| B      | Product Name  | Name of the product sold                   |
| C      | Sales Channel | Channel through which the product was sold |
| D      | Sales Amount  | Monetary value of the transaction          |
| E      | Units Sold    | Number of units sold                       |
| F      | Region        | Geographical sales region                  |

**Dataset size:** 34 records and 6 original columns.


## 1. ➕ Basic Sales Calculations — SUM, COUNT, COUNTA, and AVERAGE

### A. Calculate Total Sales

Use the SUM function to calculate the total sales amount across all transaction records.

**Excel Formula:**

```excel
=SUM(D2:D35)
```

**Explanation:** Adds all numeric values in the Sales Amount column.

**Business purpose:** Helps estimate the total sales value recorded in the dataset.

### B. Count the Number of Transactions

Use COUNTA to count non-empty transaction identifiers.

**Excel Formula:**

```excel
=COUNTA(A2:A35)
```

**Explanation:** Counts all non-empty cells in the Order Code column.

For a dataset containing 34 populated records, the expected count is 34.

Alternatively, use COUNT to count numeric sales amounts:

```excel
=COUNT(D2:D35)
```

This counts numeric entries in the Sales Amount column, not every transaction regardless of its contents.

### C. Calculate Average Sales Value

Use AVERAGE to calculate the average transaction amount.

**Excel Formula:**

```excel
=AVERAGE(D2:D35)
```

**Explanation:** Calculates the arithmetic mean of the numeric sales amounts.

**Business purpose:** Helps understand the typical value of a recorded sales transaction.

---

## 2. 📉 Identify Minimum and Maximum Sales — MIN and MAX

### A. Find the Lowest Transaction Amount

**Excel Formula:**

```excel
=MIN(D2:D35)
```

**Explanation:** Returns the smallest numeric value in the Sales Amount column.

### B. Find the Highest Transaction Amount

**Excel Formula:**

```excel
=MAX(D2:D35)
```

**Explanation:** Returns the largest numeric value in the Sales Amount column.

**Business purpose:** These functions help identify the lowest- and highest-value transactions for further review.

*Note: The actual minimum and maximum amounts depend on the sales dataset and are not specified here.*

---

## 3. 🧠 Transaction Classification Using IF

Create a new column named **Sales Category** to classify transactions according to their sales amount.

### Business Rules

- Sales Amount greater than or equal to 5,000: `High Value`
- Sales Amount below 5,000: `Regular Value`

The threshold of 5,000 is an illustrative business rule and can be adjusted according to the dataset's currency and business requirements.

### Excel Formula

Add `Sales Category` in cell G1.

Enter the following formula in cell G2:

```excel
=IF(D2>=5000,"High Value","Regular Value")
```

Copy the formula down through cell G35.

### Illustrative Example

| Sales Amount | Sales Category |
| -----------: | -------------- |
|        7,500 | High Value     |
|        2,400 | Regular Value  |
|        5,000 | High Value     |
|          850 | Regular Value  |

**Business purpose:** The IF function automatically categorizes each transaction based on a defined sales threshold.

---

## 4. 🧮 Regional Sales Analysis — SUMIF and COUNTIF

### A. SUMIF — Calculate Total Sales for the South Region

Calculate the total sales amount for transactions recorded in the South region.

**Excel Formula:**

```excel
=SUMIF(F2:F35,"South",D2:D35)
```

**Formula components:**

- `F2:F35` — Region criteria range.
- `"South"` — Region to match.
- `D2:D35` — Sales amounts to add.

**Business purpose:** Helps compare sales contributions across geographical regions.

### B. COUNTIF — Count Transactions Below 1,000

Count the number of transactions with a sales amount below 1,000.

**Excel Formula:**

```excel
=COUNTIF(D2:D35,"<1000")
```

**Explanation:** Counts numeric sales amounts that are strictly less than 1,000.

**Business purpose:** Helps measure the frequency of lower-value transactions.

*Note: Regional totals and transaction counts must be calculated from the actual data. No illustrative totals are presented as actual findings.*

---

## 5. 🔤 Extract Transaction Details — LEFT, RIGHT, and MID

The Order Code is assumed to follow this illustrative format:

`DD-MMM-BB`

Example: `28-JAN-SO`

- `DD` — Day code
- `MMM` — Month abbreviation
- `BB` — Branch code

This identifier format is a proposed structure for the alternative scenario. Use it only if the actual transaction IDs follow the same pattern.

### A. LEFT — Extract the Day

Create a new column named **Day** in cell H1.

**Excel Formula:**

```excel
=LEFT(A2,2)
```

**Illustrative example:**

| Order Code | Day |
| ---------- | --- |
| 28-JAN-SO  | 28  |
| 15-FEB-NE  | 15  |
| 03-MAR-WE  | 03  |

**Explanation:** Extracts the first two characters from the Order Code.

### B. RIGHT — Extract the Branch Code

Create a new column named **Branch Code** in cell I1.

**Excel Formula:**

```excel
=RIGHT(A2,2)
```

**Illustrative example:**

| Order Code | Branch Code |
| ---------- | ----------- |
| 28-JAN-SO  | SO          |
| 15-FEB-NE  | NE          |
| 03-MAR-WE  | WE          |

**Explanation:** Extracts the final two characters, representing the branch code in this proposed identifier format.

### C. MID — Extract the Month

Create a new column named **Month** in cell J1.

**Excel Formula:**

```excel
=MID(A2,4,3)
```

**Illustrative example:**

| Order Code | Month |
| ---------- | ----- |
| 28-JAN-SO  | JAN   |
| 15-FEB-NE  | FEB   |
| 03-MAR-WE  | MAR   |

**Explanation:** Starts at character position 4 and extracts 3 characters to return the month abbreviation.

**Business purpose:** Extracted date components and branch identifiers can help organize transaction records for reporting.

---

## 6. 📋 Final Dataset Structure

After adding the calculated columns, the dataset will contain the following fields:

| Column | Field Name     |
| ------ | -------------- |
| A      | Order Code     |
| B      | Product Name   |
| C      | Sales Channel  |
| D      | Sales Amount   |
| E      | Units Sold     |
| F      | Region         |
| G      | Sales Category |
| H      | Day            |
| I      | Branch Code    |
| J      | Month          |

Enter the formulas in row 2 and copy them down through row 35 to process all 34 records.

---

## 7. 📚 Summary of Excel Functions

| Function | Purpose                              | Formula                                      |
| -------- | ------------------------------------ | -------------------------------------------- |
| SUM      | Calculate total sales                | `=SUM(D2:D35)`                               |
| COUNT    | Count numeric sales amounts          | `=COUNT(D2:D35)`                             |
| COUNTA   | Count non-empty order codes          | `=COUNTA(A2:A35)`                            |
| AVERAGE  | Calculate average transaction value  | `=AVERAGE(D2:D35)`                           |
| MIN      | Find the lowest transaction amount   | `=MIN(D2:D35)`                               |
| MAX      | Find the highest transaction amount  | `=MAX(D2:D35)`                               |
| IF       | Classify transactions by sales value | `=IF(D2>=5000,"High Value","Regular Value")` |
| SUMIF    | Calculate South region sales         | `=SUMIF(F2:F35,"South",D2:D35)`              |
| COUNTIF  | Count transactions below 1,000       | `=COUNTIF(D2:D35,"<1000")`                   |
| LEFT     | Extract the day code                 | `=LEFT(A2,2)`                                |
| RIGHT    | Extract the branch code              | `=RIGHT(A2,2)`                               |
| MID      | Extract the month abbreviation       | `=MID(A2,4,3)`                               |

---

## 8. 📊 Business Insights

The completed analysis can help answer questions such as:

- What is the total sales value across all recorded transactions?
- How many transactions are included in the dataset?
- What is the average transaction value?
- Which transactions have the lowest and highest sales amounts?
- How much sales value comes from the South region?
- How many transactions are below 1,000?
- How many transactions fall into each sales category?
- Which months and branch codes appear in the transaction records?

**Important:** These are analysis questions, not verified findings. Actual conclusions should be added after applying the formulas to the real sales data.

---

## 9. 🎓 Key Learnings

Through this project, I practiced:

- Using SUM, COUNT, COUNTA, AVERAGE, MIN, and MAX for sales calculations.
- Applying IF to classify transaction values.
- Using SUMIF and COUNTIF for conditional analysis.
- Extracting text using LEFT, RIGHT, and MID.
- Creating calculated columns in Excel.
- Applying formulas to a practical retail sales scenario.
- Preparing data for business reporting and further analysis.

## 10. 🚀 Future Improvements

- Calculate total units sold by region.
- Compare sales performance across products and channels.
- Use sorting, filtering, and conditional formatting.
- Create PivotTables and PivotCharts.
- Build an interactive retail sales dashboard.
- Extend the analysis using SQL, Python, and Power BI.

## 👨‍💻 Author

**Subashini RK&#x20;**
Aspiring Data Analyst

This project demonstrates foundational Excel skills for data analysis and business reporting.

## ⭐ Acknowledgment

Created for educational purposes to develop practical Microsoft Excel skills and build a foundation for more advanced data analytics projects.
