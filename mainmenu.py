#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd 
#Import an Excel file and convert it to a pandas dataframe
xl = pd.ExcelFile("OSSalesData.xlsx")
lines = "="*25 + "\n"
#select a specific sheet
SalesData= xl.parse("Orders")

#function that show profit column
def SubCatProfit():#
    #accessing data from profit column
    SubCatData = SalesData[["Sub-Category", "Profit"]]
    print(lines)
    #next we are going to use functions to group by state and sum then sort
    SubCatProfit = SubCatData.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
    #going to print out the top 10 rows 
    print(SubCatProfit.head(10))

#function that shows products sales and profits
def ProductProfSales():
    #accessing data from sales column
    ProductData = SalesData[["Product Name", "Profit", "Sales"]]
    print(lines)

    ProdProfSales = ProductData.groupby(by = "Product Name").sum().sort_values(by = "Profit")
    #going to print out the top 10 rows
    print(ProdProfSales.head(10))

#function that shows unprofitable sub-category
def NegSubCat():
    ProductData = SalesData[["Sub-Category", "Product Name", "Profit"]]
    NegSubCats = ["Tables", "Bookcases", "Supplies"]
#iterate through each row and concatenate 
    for subcat in NegSubCats:
        ProductInfo = ProductData.loc[ProductData["Sub-Category"]== subcat]
        ProdProfit = ProductInfo.groupby(by = "Product Name").sum().sort_values(by = "Profit")
    print(subcat)
    print(ProdProfit)
    print(lines)

#function that shows regional effects of Sub-Cat Profits    
def SubCatProfReg():
#stroing unique value in  variable region
    regions = SalesData.Region.unique()
    print(regions)
    SubCatData = SalesData[["Sub-Category", "Profit", "Region"]]

    for region in regions:
        RegSubCatData = SubCatData.loc[SubCatData["Region"]== region]
        SubCatProfit = RegSubCatData.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
    print(lines)
    print(region)
    print(SubCatProfit.head(10))

#Customer Segment Effects of Sub-Cat Profit
def SubCatProfSeg():
#Return unique values of Series object.
    segments = SalesData.Segment.unique()
    print(segments)
    SubCatData = SalesData[["Sub-Category", "Profit", "Segment"]]


    for segment in segments:
        RegSubCatData = SubCatData.loc[SubCatData["Segment"]== segment]
        SubCatProfit = RegSubCatData.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
    print(lines)
    print(segment)
    print(SubCatProfit.head(10))

#Annual Effects of Sub-Cat Profit
def SubCatProfYear():
    SalesDataYear = SalesData
    #return the year of the order
    SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
    years= SalesData.Year.unique()

    SubCatData = SalesDataYear[["Sub-Category", "Profit", "Year"]]
# loop through the years, grouping by sub-catgory and summing up salse.
    for year in years:
        SubCatDataByYear = SubCatData.loc[SubCatData["Year"]== year]
        #remove the year column so that .sum doesn't tatal it 
        SubCatProfitNoYear = SubCatDataByYear[["Sub-Category", "Profit"]]
        SubCatProfit = SubCatData.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
        SubCatProfit = SubCatProfitNoYear.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
        print(lines)
        print(year)
        print(SubCatProfit.head(10))
#function that shows discounts 
def SubCatDiscount():
    #have data from the discount column
    SubCatData = SalesData[["Sub-Category", "Discount"]]
    print(lines)
 
    SubCatDiscount = SubCatData.groupby(by = "Sub-Category").mean().sort_values(by = "Discount")
    print(SubCatDiscount)

#menu function
def MainMenu():
#just a decoration    
    print("\n" + "*"*60)
#create menu options
    print("\nEnter (1) to see Sub-Category Profits" + 
          "\nEnter (2) to see Sub-Categories with Negative Profits" +
          "\nEnter (3) to see Sales and Profits of the 10 least profitable products" +
          "\nEnter (4) to see Sub-Categories Profits by Customer Segment" +
          "\nEnter (5) to see Sub-Categories Profits by Region" +
          "\nEnter (6) to see Sub-Categories Profits by Year" +
          "\nEnter (7) to see Sub-Categories Profits Discount Averages" +
          "\nEnter (8) to Exit")
#input statement for the user
    choice = input("Please enter a number 1 and 8: ")
#if statements that will call different functions based on the user's selection
    if choice == "1":
        SubCatProfit()
        MainMenu()
    elif choice == "2":
        NegSubCat()
        ProductProfSales()
        MainMenu()
    elif choice == "3":
       ProductProfSales()
       MainMenu()
    elif choice == "4":
        SubCatProfSeg()
        MainMenu()
    elif choice == "5":
        SubCatProfReg()
        MainMenu()
    elif choice == "6":
        SubCatProfYear()
        MainMenu()
    elif choice == "7":
        SubCatDiscount()
        MainMenu()
    elif choice == "8":
        exit()
    else:
       #Welcome message
        print("\ninvalid option. Please enter an option from 1-8 below: ")
        MainMenu()#giving the user another chance to enter an acceptable value  if he or she didn’t enter a correct value on a previous attempt
        
print("\nWelcome to the Office Solutions Sales Data Analytics System!")
MainMenu()
    
    
    
    





    

