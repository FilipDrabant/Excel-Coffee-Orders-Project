# Excel-Coffee-Orders-Project
Project in Excel transforming/cleaning Data and creating a Interactive Dashboard

Created to practice Excel skills using: XLOOKUP, INDEXing, MATCH, Pivot Tables, Pivot Charts, Formatting, Cell Formatting, Connected Timeline and Slicers...

## Goals
1. Fill out missing columns
2. Clean up Data
3. Create Dashboard

## Filling out missing columns

Dataset came with empty tables:

![image](https://github.com/user-attachments/assets/7adb19da-771d-4fca-92e5-4fe1b41e5e91)

Using XLOOKUP to fill out Customer Name and Country column from data in other sheet:
```
=XLOOKUP(C2;customers!$A$1:$A$1001;customers!$B$1:$B$1001;;0)
```
Filling out emails column left some cells with 0's as not every user has email.

Adding IF function to the statement to get rid of 0's:
```
=IF(XLOOKUP(C2;customers!$A$1:$A$1001;customers!$C$1:$C$1001;;0)=0;"";XLOOKUP(C2;customers!$A$1:$A$1001;customers!$C$1:$C$1001;;0))
```

All the information about Coffee, so columnsL Coffee Type, Roast Type and so on is in the products sheet. Since Columns in product sheet are In same order we can use INDEX and MATCH functions to fill our table:
```
=INDEX(products!$A$1:$G$49;MATCH(orders!$D2;products!$A$1:$A$49;0);MATCH(orders!J$1;products!$A$1:$G$1;0))
```
Sales column filled out by simple multiplication function:
```
=L2*E2
```
Filled out table:

![image](https://github.com/user-attachments/assets/4052a033-8cc7-4f2e-a405-fef22bbb5cb1)

## Cleaning up Data
First we should provide full names from abbreviations like Coffe Type: Rob to Robusta and Roast Type: L to Light and so on.

We can do that by creating new columns using Multiple IF functions:
```
=IF(J2="M";"Medium";IF(J2="L";"Light";IF(J2="D";"Dark";"")))
```
Result:

![image](https://github.com/user-attachments/assets/e713ac8d-ca2d-445d-94a2-583c8a00b1d2)

Next is transforming the date types to dd/mm/yy format to dd-mmm-yy format which looks cleaner.

Do that using Custom Cell formatting:

![image](https://github.com/user-attachments/assets/77d8cb04-4c29-4c96-b5d0-79346dc5a960)

Also using Cell formatting adding unit type and currency:

![image](https://github.com/user-attachments/assets/dae02527-dcce-4453-949f-647f36a693ca)

Also checked the table for duplicates using Remove Duplicates but this table didn't have any.

## Creating Dashboard
To create a Dashboard using we need only specific data for each element of our dashboard. To do that best way is to create pivot tables from main sheet.

First we need to convert our sheet into a Table using ctrl+T and then having selected any cell in our table press ALT+N+V+T to create Pivot Table on new sheet.

On a new sheet we can now select a data we want and create elements:

![image](https://github.com/user-attachments/assets/4c951470-5231-47c3-86cc-04475177b225)

Sales by country Chart created from pivot table usign custom 'purple' formatting:

![image](https://github.com/user-attachments/assets/9e400514-5fd7-492f-977a-d08795673a29)

Final Interactive Dashboard by combing 3 pivot tables and creating custom Purple  formatting for each element:

![image](https://github.com/user-attachments/assets/33c2e65f-d31a-46e6-a4b1-374794f48b68)










