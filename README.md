# Cofee Sale Data Analysis (using Excel)
## Introduction:
This report presents a data analysis conducted on a dataset containing information about cofee sales. 

The dataset
<a href="https://github.com/ThinhThuong/Excel-Project---Coffee-Sales/blob/main/Project%20Coffee%20Sales.xlsx">Coffee Sales</a>
include 3 sheets:
- Sheet *customers*: customer information
- Sheet *products*: coffee information
- Sheet *order*: compile customer orders and calculate sales. Ensure that some columns are filled with the required data before proceeding with the sales calculations

![image](https://github.com/user-attachments/assets/3368abec-921d-4d10-b3ba-9df6be664fe2)

## Objective:
Analyze coffee sales data across multiple years to identify trends, patterns, and key insights for business performance and decision-making

## Question:
1. Which country recorded the highest coffee sales?
2. What is the annual coffee sales performance for each year?
3. Who are the top 5 customers with the highest coffee purchases in 2022?

## Process:
### 1. Fill Data
Start populating columns from `Customers Name` (Column F) to `Sales` (Column M) with the following instructions:

* Column `Customers Name` and `Country`: use the XLOOKUP function, where the lookup_array and return_array are sourced from the *customers* sheet.

* Column `Email`: to handle blank values, use a combination of IF and XLOOKUP. Ensure that if the returned value is blank, the result is displayed as a blank value ("") instead of 0.

![image](https://github.com/user-attachments/assets/2fa92484-5639-42af-9d6f-0f295a4aa4b0)

* From column `Coffee Type` to `Unit Price`: use INDEX function instead of XLOOKUP to fill value from *products* sheet.

![image](https://github.com/user-attachments/assets/acbc439f-7af5-4f0b-b26f-53309e745d93)

* Column `Sales`: calculate Sales = Unit Price * Size
* Column `Loyalty Card` is added: use the XLOOKUP function, where the lookup_array and return_array are sourced from the *customers* sheet.

### 2. Simplify Confusing Values
Replace ambiguous values in the columns `Coffee Type` and `Roast Type` for better readability:
- For `Coffee Type`: Ara => Arabica, Exc => Excelsa, Rob => Robusta, Lib => Liberica
- For `Roast Type`: L => Light, M => Medium, D => Dark

### 3. Format Data Type
- Format date data: for column `Order Date`, convert the original format (e.g., 05-09-19) to a clear format with the dd-mmm-yyyy style.
For example, 05-09-19 becomes 05-Sep-2019.
- Format decimal places and add unit:
  * Column `Size`: adjust values to 1 decimal place and append the unit *kg*
  * Column `Unit Price` and `Sales`: adjust values to 2 decimal places and append the unit *$*
 
### 4. Eliminate Duplicate Records

Begin by removing duplicate rows to ensure the dataset is clean and accurate for analysis.

Once duplicates are removed, the final dataset will be prepared and formatted for data visualization, as illustrated in the image below.

![image](https://github.com/user-attachments/assets/4c5082d8-a1bb-4c7b-a3f2-086e32db6a05)

### 5. Create Pivot Tables for Insights
- Dedicate a separate sheet for pivot tables to facilitate easy updates and adjustments.
- Build pivot tables based on specific questions and outline the key metrics.
- Lay the groundwork for a simple dashboard to present data interactively.

![image](https://github.com/user-attachments/assets/5b347ffd-c318-45a9-9f2f-af859fc788b2)

### 6. Design an Interactive Dashboard
- Create a dedicated sheet for the dashboard.
- Incorporate slicers to filter variables dynamically and uncover deeper insights.
- Format the dashboard and slicers with a cohesive and professional theme to enhance visual appeal.

<a href="https://github.com/ThinhThuong/Excel-Project---Coffee-Sales/blob/main/Interactive%20dashboard%20-%20Coffee%20Sales.gif">View interactive dashboard</a>

## Answer the questions:
**1. Which country recorded the highest coffee sales?**

![image](https://github.com/user-attachments/assets/2f0975a5-f386-4176-9684-1091d3bcbcac)

From 2019 to 2022, the United States led global coffee sales with $35,639, a figure five times higher than Ireland's and twelve times higher than the United Kingdom's.

**2. What is the annual coffee sales performance for each year?**

![image](https://github.com/user-attachments/assets/cd18e618-36ac-4425-b0de-7ebcce7f9a4d)

- During the period from 2019 to 2022, coffee sales showed an overall downward trend.

- Robusta was the least popular coffee type, consistently recording the lowest sales compared to other types.

- In 2019 and 2020, Exelsa had the highest sales, exceeding $3,400.

- In 2021, Arabica became the most consumed coffee type, with sales reaching $4,046.

- By 2022, coffee consumption declined significantly, with sales for all coffee types dropping to nearly half of the previous year's figures.

**3. Who are the top 5 customers with the highest coffee purchases in 2022?**

![image](https://github.com/user-attachments/assets/a803b2ba-5900-47fd-bcbf-e54acee68b84)

- In 2022, the United Kingdom had no representation in the top 5 coffee buyers: the top 4 customers were from the United States, while the 5th spot was held by a customer from Ireland.
- The coffee purchases of the top 5 customers ranged from $164 to $205.

## Recommendations:

**1. Focus on Emerging Markets:**

- Target countries with untapped potential, such as the United Kingdom, where representation in top coffee buyers is currently absent.
- Conduct market research to identify and address barriers to coffee consumption in these regions.

**2. Revitalize Coffee Consumption Across All Types:**

- Launch promotional campaigns or limited-time offers to increase sales of underperforming coffee types like Robusta.
- Highlight health benefits, unique flavors, or cultural associations of different coffee types to attract a diverse consumer base.

**3. Leverage Top-performing Coffee Types:**

- Build marketing campaigns around Exelsa and Arabica, focusing on their historical popularity and consumer appeal.
- Develop new products or blends using these coffee types to spark renewed interest.

**4. Expand and Incentivize High-value Customers:**

- Introduce loyalty programs and discounts targeted at high-value customers to encourage repeat purchases.
- Use personalized marketing to cater to the preferences of top buyers, particularly those in the United States and Ireland.
