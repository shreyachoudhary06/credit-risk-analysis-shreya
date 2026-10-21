# Build Your Credit Risk Analyst Project

This is a hands-on learning path for rebuilding the project from its raw data. Work through one phase at a time, save your own outputs, and use the supplied finished files only to compare your results.

## What you will build

You will create a cleaned loan dataset in Excel, explore it with MySQL, and build a five-page Power BI report. The repository already contains the finished dashboard, workbook, and cleaned CSV as reference material; your practice copies should not overwrite them.

The data has one row per loan application. `loan_status` is the outcome: `0` means the loan did not default and `1` means it defaulted. Your main business question is: **Which borrower and loan characteristics are associated with a higher default rate?**

## Tools and files

- **Excel** — cleaning, formulas, and quality checks.
- **MySQL Workbench and MySQL Server** — database tables and analysis queries.
- **Power BI Desktop** — data model, measures, and interactive report.
- Raw input: `data\credit_risk_dataset.csv`
- Finished comparison files: `data\credit_risk_cleaned.csv`, `excel\credit_risk_cleaned.xlsx`, and `Credit_Risk_Dashboard.pbix`
- Practice SQL: `sql\beginner_workshop.sql`

No code packages are required. MySQL Workbench connects to a MySQL Server; the server's local `root` password is whatever was chosen at installation. Do not put a password in this repository.

## Phase 0 — Make a safe practice copy

1. In File Explorer, open the repository folder and make a separate folder outside it called `credit-risk-practice`.
2. Copy `data\credit_risk_dataset.csv` into that folder. Leave the repository's raw and cleaned data, Excel workbook, and PBIX unchanged.
3. In Excel, open your copied raw CSV. Immediately use **File → Save As** and save it as `credit_risk_practice.xlsx` in your practice folder. Choose the `.xlsx` format so your cleaning work can have multiple sheets and formulas.
4. In the Excel workbook, rename the original sheet `Raw_Data`. Duplicate it and name the copy `Cleaned_Data`. Do all edits on `Cleaned_Data`.

**Learn:** Raw data is an evidence-preserving source. Analysts keep an untouched copy so they can explain, reproduce, and correct transformations later.

## Phase 1 — Inspect and clean in Excel

### 1. Understand the data before changing it

1. On `Raw_Data`, note the number of data rows, the 12 column names, and what one row represents.
2. Turn the data on `Cleaned_Data` into an Excel Table with **Ctrl+T** and confirm **My table has headers**.
3. Use the filter arrows and **Data → Sort & Filter** to inspect values in `person_age`, `person_emp_length`, `loan_int_rate`, and `loan_status`.
4. Record row counts before and after every cleaning operation in a new `Cleaning_Log` sheet. Add columns for `Step`, `What I checked`, `Action`, and `Rows before/after`.

**Learn:** Always check the unit of observation, column meaning, missingness, and impossible values before calculating metrics.

### 2. Remove exact duplicate records

1. Select a cell in the `Cleaned_Data` table and choose **Data → Remove Duplicates**.
2. Keep all 12 original columns selected, then confirm.
3. Record how many duplicates Excel removed. The raw file is expected to contain 165 exact duplicate rows.

**Learn:** A duplicate is defined relative to the chosen columns. Removing duplicates on only one column, such as age, would incorrectly discard different borrowers or loans.

### 3. Remove impossible values

1. Filter `person_age` to values above 80. Inspect the rows, then remove the seven rows with impossible ages documented by this project.
2. Filter `person_emp_length` to `123` and remove the two impossible records.
3. Clear the filters and record the row count.

**Learn:** Filtering and inspecting before deletion is safer than applying a broad rule without seeing which records it affects.

### 4. Fill missing numeric values with the median

The raw CSV has 3,116 missing interest rates and 895 missing employment lengths. The README reports 887 missing employment lengths during its cleaning stage; counts can differ when invalid rows have already been removed. Calculate the count from the data at the stage you are cleaning rather than assuming either number.

For each field:

1. In a spare cell, calculate its median with `=MEDIAN(Cleaned_Data[loan_int_rate])` or `=MEDIAN(Cleaned_Data[person_emp_length])`. Write the result in `Cleaning_Log`.
2. Add a temporary helper column. For interest rate, use:

   ```excel
   =IF([@loan_int_rate]="",MEDIAN(Cleaned_Data[loan_int_rate]),[@loan_int_rate])
   ```

   For employment length, use the same pattern with `person_emp_length`.
3. Fill the formula down. Copy the helper column and use **Paste Special → Values** over the original column.
4. Remove the temporary helper column and check that the blank count is zero.

**Learn:** The median is less affected by extreme values than the mean. Imputation is a modeling choice; record it, and do not pretend the estimated values were observed.

### 5. Add three feature columns

Add these columns to the right of the original 12 fields:

| Column | Excel Table formula | Meaning |
|---|---|---|
| `income_outlier_flag` | `=--([@person_income]>1000000)` | 1 if annual income exceeds 1,000,000; otherwise 0. Flag rather than delete valid high incomes. |
| `loan_status_label` | `=IF([@loan_status]=1,"Defaulted","Not Defaulted")` | Human-readable version of the target. |
| `risk_category` | `=IF([@loan_grade]="A","Low Risk",IF(OR([@loan_grade]="B",[@loan_grade]="C"),"Medium Risk","High Risk"))` | Groups A as low, B/C as medium, and D–G as high risk. |

**Learn:** Feature engineering creates useful analysis fields while preserving the original variables.

### 6. Validate and save the Excel output

Check that:

- The cleaned table has 32,407 rows and 15 columns.
- `loan_status` contains only 0 and 1; `risk_category` contains Low, Medium, or High Risk.
- `loan_int_rate` and `person_emp_length` have no blanks after imputation.
- The `Cleaning_Log` explains every deletion and fill operation.

Save the workbook. Then select `Cleaned_Data` and use **File → Save a Copy / Save As → CSV UTF-8** to create `credit_risk_practice_cleaned.csv` in your practice folder. CSV stores only the active sheet, so keep the `.xlsx` workbook as your full record.

**Compare:** The supplied `data\credit_risk_cleaned.csv` has the expected 32,407 rows and 15 columns. Compare it with your output, but investigate differences instead of silently changing your work to match.

## Phase 2 — Import and analyze with MySQL

### 1. Create a separate practice schema

1. Start MySQL Workbench and connect to your local MySQL Server.
2. Open `sql\beginner_workshop.sql` from the repository.
3. Read the first statements. They create a new schema called `credit_risk_learning` and an empty table called `credit_risk_workshop`; they do not drop or replace existing schemas or tables.
4. Run the setup statements. You can select a statement and press **Ctrl+Enter** to run just that statement.

**Learn:** A schema is a container for tables. Keeping practice work in its own schema avoids altering the portfolio database or other databases.

### 2. Load your cleaned CSV

1. In the Navigator, refresh `credit_risk_learning`, expand **Tables**, then right-click `credit_risk_workshop`.
2. Choose **Table Data Import Wizard**, select your practice cleaned CSV, and choose the existing `credit_risk_workshop` table as the destination.
3. Confirm that the CSV headers map to the same-named columns and review the preview before starting the import.
4. Run the verification query at the top of the SQL file. It should show 32,407 rows. Do not import the same file again into the same table; repeated imports append duplicate data.

The CSV provided in the repository can be used for a first pass if you have not completed the Excel phase yet. Prefer importing your own Excel output when practicing end to end.

### 3. Explore the analysis queries

Run the queries in `sql\beginner_workshop.sql` one at a time. Before each query, write down what you expect; after it runs, record one finding and one limitation. The queries cover:

- Overall default rate and row counts.
- Default rates by grade, intent, home ownership, age group, prior default, and risk category.
- Average borrower/loan characteristics by outcome.
- Compound segments such as grade + intent + ownership.
- Default rates by credit-history length.

**Learn:** `GROUP BY` defines the population for each result row. A percentage should use defaults divided by all loans in that group—not the whole table. Always show the group size alongside a rate; a high rate from a tiny group may be unstable.

## Phase 3 — Build the report in Power BI Desktop

### 1. Load and inspect the data

1. Open Power BI Desktop and choose **Get data → Text/CSV**.
2. Select your cleaned CSV. Choose **Transform Data** so you can inspect it before loading.
3. In Power Query, confirm the preview has 15 fields, rename the query `credit_risk_cleaned`, and set sensible types: whole number for age, income, loan amount, status, and history length; decimal number for employment length, interest rate, and loan-to-income; text for categories.
4. Check for errors or unexpected blank cells, then choose **Close & Apply**.

**Learn:** Power Query handles repeatable import and transformation; the model stores the loaded table; DAX measures calculate results in response to report filters.

### 2. Create two useful segmentation columns

In Power Query, choose **Add Column → Custom Column**:

**`Age_Group`**

```powerquery
if [person_age] <= 25 then "20-25"
else if [person_age] <= 30 then "26-30"
else if [person_age] <= 40 then "31-40"
else if [person_age] <= 55 then "41-55"
else "55+"
```

**`LTI_Bucket`**

```powerquery
if [loan_percent_income] < 0.10 then "< 10%"
else if [loan_percent_income] < 0.20 then "10-20%"
else if [loan_percent_income] < 0.30 then "20-30%"
else if [loan_percent_income] < 0.40 then "30-40%"
else "> 40%"
```

Choose appropriate text data types, then **Close & Apply**.

### 3. Create the core DAX measures

Select the `credit_risk_cleaned` table and choose **New measure**. Create these measures one by one:

```dax
Total Applicants =
COUNTROWS(credit_risk_cleaned)

Total Defaulters =
CALCULATE(
    [Total Applicants],
    credit_risk_cleaned[loan_status] = 1
)

Non Defaulters =
[Total Applicants] - [Total Defaulters]

Default Rate =
DIVIDE([Total Defaulters], [Total Applicants])

Loan Amount at Risk =
CALCULATE(
    SUM(credit_risk_cleaned[loan_amnt]),
    credit_risk_cleaned[loan_status] = 1
)

Average Interest Rate =
AVERAGE(credit_risk_cleaned[loan_int_rate])

Average Loan-to-Income =
AVERAGE(credit_risk_cleaned[loan_percent_income])
```

Select `Default Rate` and set its format to **Percentage** with two decimal places. Do not multiply the measure by 100 when it is percentage-formatted.

**Learn:** A measure recalculates under slicers and visual filters. A calculated column is stored per row; use it for labels/buckets and measures for aggregations.

### 4. Build page 1 before attempting all five pages

Create a page named **Executive Summary**:

1. Add Card visuals for `Total Applicants`, `Total Defaulters`, `Default Rate`, and `Loan Amount at Risk`.
2. Add a clustered column chart: `loan_grade` on the X-axis, `Default Rate` on the Y-axis.
3. Add a bar chart: `loan_intent` on the Y-axis, `Default Rate` on the X-axis.
4. Add a slicer for `person_home_ownership` and another for `risk_category`.
5. Add a clear title, readable number formats, and descriptive chart titles.
6. Save As `Credit_Risk_Dashboard_Practice.pbix` in your practice folder.

**Checkpoint:** The full-data cards should show 32,407 applicants, 7,088 defaulters, and a 21.87% default rate (allow for rounding). Select a slicer and confirm the charts and measures change.

### 5. Extend it into the five-page replica

Build one page at a time; test filters and drill behavior before moving on.

| Page | First visuals to create | Question answered |
|---|---|---|
| Executive Summary | KPI cards, default rate by grade and intent, slicers | How healthy is the portfolio? |
| Borrower Profile | Default rate by age group, home ownership, prior-default status | Which borrower segments have higher observed default rates? |
| Grade Deep Dive | Grade-level KPI and intent-by-grade chart; add drill-through on `loan_grade` | What is happening inside a selected grade? |
| Financial Impact | Loan amount at risk by grade, average interest rate by grade, default rate by LTI bucket | Where is exposure concentrated and how does it relate to loan-to-income? |
| Scorecard | Text cards comparing selected high- and lower-risk groups, plus key findings | What should a stakeholder remember? |

Start with standard Power BI visuals and the measures above. Add custom tooltips, a What-If parameter, decomposition tree, and theme only after the core pages work. Those are optional advanced exercises, not prerequisites.

## Phase 4 — Present and explain your work

1. Save a final practice PBIX and Excel workbook in your practice folder.
2. Write a short summary: one finding, the evidence behind it, a limitation, and a business action to investigate.
3. Compare your report with `Credit_Risk_Dashboard.pbix` and screenshots in `images\`. Note design or analysis differences and why you made them.
4. Keep the original files intact. If you want to publish the portfolio later, verify the dataset license, remove personal credentials, and review any potentially sensitive material first.

## Common beginner questions

- **Workbench cannot connect:** confirm the MySQL80 service is running, use the same local host/port selected during installation, and enter the password you created. Do not paste credentials into SQL or commit them.
- **Import fails due to column types:** check the import preview and destination mapping. The workshop table uses numeric types for numeric fields; make sure decimal fields use a dot (`.`) as the decimal separator.
- **Your row count differs:** check whether you removed duplicates on all original columns, applied the same invalid-value rules, and imported the file once. Record the stage and investigate rather than hiding the mismatch.
- **A Power BI number looks 100 times too large:** check whether a ratio measure was multiplied by 100 and also formatted as Percentage.
- **A chart shows “Count of loan_status”:** use the `Default Rate` measure as the value, not the implicit count or sum of the status column.

## Vocabulary to keep

| Term | Plain-English meaning |
|---|---|
| Default rate | Defaults divided by all loans in the group being analyzed. |
| Grain | What one row represents; here, one loan application. |
| Measure | A calculation that changes with filters in a Power BI report. |
| Dimension | A descriptive grouping such as grade, intent, or home ownership. |
| Data quality check | A test that confirms values and row counts are plausible before analysis. |
| Limitation | A reason an observed pattern should not automatically be treated as causation or a lending rule. |
