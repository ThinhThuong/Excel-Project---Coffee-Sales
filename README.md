# Cofee Sale Data Analysis (using Excel)
## Introduction:
This report presents a data analysis conducted on a dataset containing information about cofee sales. 
## Dataset: 
The dataset

<a href="https://github.com/ThinhThuong/Excel-Project---Coffee-Sales/blob/main/Project%20Coffee%20Sales.xlsx">Coffee Sale</a>
include 3 sheets:
- Sheet `customers`: customer information
- Sheet `products`: coffee information
- Sheet `order`: compile customer orders and calculate sales. Ensure that some columns are filled with the required data before proceeding with the sales calculations

![image](https://github.com/user-attachments/assets/790db195-336a-4b10-be8a-fd516be614d6)

## Objective:
Analyze coffee sales data across multiple years to identify trends, patterns, and key insights for business performance and decision-making

## Question:
1. Which region has the highest bike sales?
2. How does average income affect the decision to buy a bike?
3. Is there a relationship between commute distance and bike purchases?
4. Which age range has the highest number of bike buyers?

## Process:
### 1. Eliminate Duplicate Records
Start by removing duplicate rows to ensure the dataset is clean and accurate for analysis.

### 2. Simplify Confusing Values
Replace ambiguous values in the columns “Marital Status” and “Gender” for better readability:
- Convert M to Married and S to Single.
- Convert F to Female and M to Male.
Example: Change M in Marital Status to Married for clarity.

### 3. Group Age into Categories
Simplify the continuous Age feature by grouping it into meaningful categories:
- Adolescent (<31 years)
- Middle Age (31–54 years)
- Old (>54 years)
- Invalid (if age falls outside these ranges)

Add a new column, “Age Bracket”, to store these categories for enhanced interpretability.

![image](https://github.com/user-attachments/assets/da986eb8-b62b-4ab3-b839-2f1638f51f5a)

### 4. Create Pivot Tables for Insights
- Dedicate a separate sheet for pivot tables to facilitate easy updates and adjustments.
- Build pivot tables based on specific questions and outline the key metrics.
- Lay the groundwork for a simple dashboard to present data interactively.

![image](https://github.com/user-attachments/assets/b425a221-16d0-4f53-971c-34a2a7fd3fd7)

### 5. Design an Interactive Dashboard
- Create a dedicated sheet for the dashboard.
- Incorporate slicers to filter variables dynamically and uncover deeper insights.
- Format the dashboard and slicers with a cohesive and professional theme to enhance visual appeal.

<a href="https://github.com/ThinhThuong/Excel-project/blob/main/Interactive%20Dashboard.gif">View interactive dashboard</a>

## Answer the questions:

![image](https://github.com/user-attachments/assets/3b2a2e04-b396-48eb-9335-4cf27eb9fc7e)
**1. Which region has the highest bike sales?**
- North America has the highest bike sales, followed by Europe and the Pacific.
- In the Pacific and North America, males tend to buy more bikes than females, whereas in Europe, women purchase more bikes than men.

**2. How does average income affect the decision to buy a bike?**
- People who purchase bikes tend to have a significantly higher income than those who do not buy bikes, for both females and males.
- The average income for female bike buyers is nearly $55,700, while the average income for male bike buyers is about $60,000.

**3. Is there a relationship between commute distance and bike purchases?**
- Most people who buy bikes have a commute distance of 0-1 miles.
- People with a commute distance of 5 miles or more tend to avoid purchasing bikes.

**4. Which age range has the highest number of bike buyers?**
- The middle-aged group (31-54 years) has a significantly higher number of bike buyers compared to non-buyers.
- In other age groups (adolescents and older adults), the number of people who do not buy bikes is higher than those who do.

## Recommendations:
- Target regions with lower sales potential: Given the high sales in North America, explore strategies to increase bike sales in regions like Europe and the Pacific, where there are gender-based purchasing differences.
- Tailor marketing to income segments: Since higher income correlates with bike purchases, focus marketing efforts on individuals with disposable income. Highlight bikes as a practical and efficient investment, particularly in affluent areas.
- Promote convenience for short commutes: Emphasize the benefits of bikes for shorter commutes, especially in urban areas, where biking can be a more efficient and eco-friendly alternative to cars.
Develop products and services for middle-aged buyers: As middle-aged individuals are the largest group of bike buyers, consider offering bikes and related products that cater to their lifestyle, including comfortable designs, easy maintenance, and convenience for commuting.
