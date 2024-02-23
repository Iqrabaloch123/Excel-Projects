# Excel-Projects
## Finance Report: Personal Expense Tracker


![111](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/1a933118-a8d5-402f-a031-34e006554335)



- **Project Objective:**
- ** Create Personal Expense Tracker(https://docs.google.com/spreadsheets/d/1ER9C1-gE_60WlVCoAize6s0ZopcV9Pgl/edit#gid=79930846)


  **1.** Develop a Personal Expense Tracker in Excel.

  **2.** Track monthly income and expenses with separate tables.

- **Purpose:**

  - Evaluate monthly financial performance.
  - Support budgeting and decision-making.
  - Facilitate communication on financial goals.

- **Key Features:**

  - Separate income and expense tables.
  - Excel's table total feature for quick calculations.
  - Conditional formatting for savings below the target.
  - Flexibility to add extra items for detailed tracking.

- **Usage Instructions:**

  - Regularly update income and expenses.
  - Monitor and strive to meet the savings target.

## Skills:

- [x] Excel proficiency for table creation.
- [x] Use of conditional formatting for visualization.
- [x] Incorporation of additional income and expenses.
- [x] Regular tracking of real expenses.

## Soft Skills:

- [x] Understanding of personal finance.
- [x] Designing a user-friendly expense tracker.
- [x] Optimizing for efficient financial management.
- [x] Systematic approach to updating and maintenance.

## product sales report:


![222](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/115f28ba-c355-4cf5-92c2-4479088eda7f)



- **Project objective:** 

    **1.** Create a product sales report(https://docs.google.com/spreadsheets/d/1f-q-vPzCiQ75BGfcWdkyL_pwyZ19Wxyk/edit?usp=drive_link&ouid=112109007698660220437&rtpof=true&sd=true)

  
(https://drive.google.com/file/d/1SnvjH1FhmjllirzG5LiNUO-itmwN7Lk8/view?usp=drive_link)


  Certainly! Here's a concise summary of the cleaning and merging tasks:

**Task 1: Cleaning the Bad Data:**

*For "orders" table:*
1. Eliminate duplicate 'order_id': `=IF(COUNTIF($A$2:A2,A2)>1,"",A2)`
2. Convert 'product_id' to Text: Change data type in Excel.
3. Replace 'qty' entries ending with "Q": `=IF(ISNUMBER(FIND("Q",B2)),SUBSTITUTE(B2,"Q",""),B2)`
4. Substitute empty 'qty' values with 'Not Available': `=IF(ISBLANK(B2),"Not Available",B2)`

*For "products" table:*
1. Remove extra spaces in 'product_name': `=TRIM(A2)`
2. Split 'price (in Rs)' at '₹' and extract numerical values: `=VALUE(TRIM(MID(A2,FIND("₹",A2)+1,LEN(A2))))`
   (Rename columns afterward)

*For "customer" table:*
1. Convert customer names to lowercase: `=LOWER(A2)`
2. Paste names permanently using 'Paste Special' -> 'Values'.

**Task 2: Merging Data:**
1. VLOOKUP for customer names based on 'customer_id': `=VLOOKUP(D2, customer!A:B, 2, FALSE)`
2. INDEX-MATCH for product names based on 'product_id': `=INDEX(products!B:B,MATCH(E2, products!A:A,0))`
3. XLOOKUP for 'price (in Rs)' based on 'product_id': `=XLOOKUP(F2,products!A:A,products!C:C)`
4. Create "total_price" by multiplying 'qty' and 'price (in INR)': `=G2*H2`

  **Loan Repayment Report:**


  ![333](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/c0384748-7acc-4f8d-a633-3cc523afdb90)


  create loan repayment report(https://drive.google.com/file/d/1ja2KzOZdZZdFD1kAgu0bnpEYuO6l2o8E/view?usp=drive_link)

  
  (https://docs.google.com/spreadsheets/d/1-i5TjyEdwSY2xB5PCDCC0TkJEuzBZgNb/edit?usp=drive_link&ouid=112109007698660220437&rtpof=true&sd=true)
  

After performing the calculations for 'Total Interest Amount' and 'Total Cost of Loan' across various 'Annual Interest Rates' and 'Loan Periods in years,' the insights provide valuable information for **Mr. Hathodawala's** decision-making.

Considering **Mr. Hathodawala's** monthly budget for **loan repayment**, set at ₹25,000, it is crucial to align the chosen loan with his financial capacity.

Review the values obtained from the calculations and cross-verify them with the figures in the "Loan_repayment_template.pdf" file to ensure accuracy and reliability.

Based on the insights derived from the PMT function, evaluating the '**Monthly Payment (EMI),' 'Total Cost of Loan,'** and '**Total Interest Amount,'** recommend the bank that offers terms most favorable to **Mr. Hathodawala's** budget and financial goals.

This comprehensive analysis will empower **Mr. Hathodawala** to make an informed decision, selecting a **loan** offer that not only fits his budget but also minimizes the overall cost of borrowing.

## power query pdf


![555](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/13e064fd-15d7-4aee-a18c-03dc22e7b624)

create power query pdf(https://docs.google.com/spreadsheets/d/1l2P_ve_McNsDrX9ovzfnzJX9rHcOa97z/edit?usp=drive_link&ouid=112109007698660220437&rtpof=true&sd=true)

(https://drive.google.com/file/d/1z9Lie3nnVrzCqDn9LLAXhX673wfjaOUV/view?usp=drive_link)

## Power Query Execution Steps:
Open a new Excel file and load the provided CSV files, "bookings_data.csv" and "rooms_data.csv," using the "From Text/CSV" option. Then, open Power Query.

Change the data type of the "property_id" column to "text."

Replace values in the "property_name" column from "Atliq bay" to "Atliq Bay."

Format the "property_type" column by removing unnecessary leading or trailing spaces using the TRIM() function.

Split the "city|city_code" column into two separate columns and rename them accordingly.

Create a new conditional column, "Availability Status," based on the conditions related to "successful_bookings" and "capacity."

Create a new custom column, "occ%," representing the ratio of successful bookings to capacity. Change the data type to a percentage format.

Merge the "bookings_data" and "rooms_data" tables on the "room_id" column. Reorder columns with "room_class" next to "room_id."

Extract the month_name from the "date" column.

Execute the steps and transformations to achieve a refined dataset for further analysis.

## report
![666](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/c2ec8772-3483-44b3-8b6f-e709f9b8a05c)


(https://drive.google.com/file/d/1db_mQ6GanQ7oaZFDQc7dVyU6Opqav7db/view?usp=drive_link) 

## report1

![777](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/867d8b21-a16d-414f-ae06-b8b35b458aa4)


(https://drive.google.com/file/d/1orPXttz5pZtHM5FJISSxcuJkdqCHrOq-/view?usp=drive_link)

## report 2

![888](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/eaab07ab-f504-42ea-8c98-046afc2bfd79)

(https://drive.google.com/file/d/15s3OVfZEhHR9DpUdwT6CRmOegAq1c60N/view?usp=drive_link)

## DAX Measures and Pivot Table:
Total Revenue in June for Business Category:


=CALCULATE(SUM(fact_bookings[revenue]), dim_properties[property_category] = "Business", dim_properties[property_name])
Most Effective Booking Platform for Atliq Grands in Week 'W 27':


=CALCULATE(
    VALUES(fact_bookings[booking_platform]),
    dim_properties[property_name] = "Atliq Grands",
    fact_bookings[week_no] = "W 27"
)
Average Rating of 'Atliq Blu' in July:


=CALCULATE(AVERAGE(fact_bookings[ratings_given]), dim_properties[property_name] = "Atliq Blu", dim_properties[month] = "July")
After creating the DAX measures and constructing the PivotTable as per the format in "report1.pdf" and "report2.pdf," you can use these insights to answer specific questions and make informed decisions related to the hospitality challenge.

## Sales Market Report

![888](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/cd994a9d-8b8b-4e24-909d-d175dc266e28)


create Sales Market Report(https://drive.google.com/file/d/1U3D1p50UxlYSGLb94iGZEtapbinA3dRx/view?usp=drive_link)

(https://docs.google.com/spreadsheets/d/1KgOqj9_5AWuMIMrUoNK0pSf5me4Z8exY/edit?usp=drive_link&ouid=112109007698660220437&rtpof=true&sd=true)

** Top 10 Products Based on Percentage Increase in Net Sales (2020 to 2021):
Utilize appropriate calculations and sort products based on percentage increase.

Division Report (Net Sales Data for 2020 and 2021 with Growth Percentage):
Present net sales data for each division.
Include growth percentage for each division.

** Top 5 and Bottom 5 Products by Quantity Sold:
Rank products based on quantity sold.
Identify the top 5 and bottom 5 products.
New Products Introduced in 2021:
Identify rows with a percentage value of 0% in the "21 vs 20" column.

** Top 5 Countries by Net Sales in 2021:
Analyze and rank countries based on net sales in 2021.
By addressing these business inquiries and presenting the findings in a structured format, this report aims to support informed decision-making and provide valuable insights into the company's sales performance and market dynamics.

## Project Priority Matrix Beautification:

![999](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/6e652e8c-c097-4141-aad2-7e1d89c913c5)



create Project Priority Matrix Beautification(https://drive.google.com/file/d/1OrXa1hu64kMJymtcKk-SP7QRQH0jBzCg/view?usp=drive_link)


(https://docs.google.com/spreadsheets/d/1-rEU9TY9z06lo8CrjWCBuoRZCq3-2dAT/edit?usp=drive_link&ouid=112109007698660220437&rtpof=true&sd=true)

To achieve the visually appealing version of the "Project Priority Matrix," modifications were made to enhance clarity and presentation. Taking inspiration from the provided "Project Priority Matrix_Beautified.pdf," the following improvements were implemented:

Adjusted color schemes for better visual appeal.
Enhanced font styles and sizes for readability.
Aligned and organized project priority categories.
Utilized graphical elements to highlight critical information.
Improved overall layout and formatting.
The goal is to provide a more engaging and comprehensible representation of project priorities, aligning with the aesthetic and functional standards set by the "Project Priority Matrix_Beautified.pdf."


## Scenario Planning Tool Beautification:

![9090](https://github.com/Iqrabaloch123/Excel-Projects/assets/130351579/75d9000f-803a-4ae8-b592-de5176c276b6)


create Scenario Planning Tool Beautification(https://drive.google.com/file/d/1ksSDQ53wChGTF3AGokgt1voGDbTZnAl2/view?usp=drive_link)

(https://docs.google.com/spreadsheets/d/15IL33EPur_G4jtHpCyVvXcWWHTP1j1Ea/edit?usp=drive_link&ouid=112109007698660220437&rtpof=true&sd=true)

To enhance the visual appeal and functionality of the "scenario planning tool.xlsx," modifications were made, drawing inspiration from the aesthetically pleasing "Scenario Planning Tool_Beautified.pdf" created by Peter. Key improvements include:

Enhanced color schemes for better visual distinction.
Improved font styles and sizes for readability.
Streamlined layout to provide a clearer structure.
Utilized graphical elements for emphasis and clarity.
Adjusted formatting for a polished and professional appearance.
The goal is to create a visually appealing version of the "scenario planning tool.xlsx" while ensuring alignment with the standards set by the "Scenario Planning Tool_Beautified.pdf."















   
   
   



