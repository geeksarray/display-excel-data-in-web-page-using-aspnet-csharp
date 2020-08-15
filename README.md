# Display Excel data in web page using ASP.NET C#

This article will describe to you how to show Microsoft Excel data to the ASP.NET Web page. This tutorial will read from excel which has data from the Northwind database's Products table. This excel sheet has multiple sheets named with Product Category Name. And each sheet will have products depending on the selected category.

All the sheet names will be shown in ASP.NET DropDownList control and on selection change of DropDownList value corresponding Products will display in GridView control.

## Files

1. [Products.xlsx](https://github.com/geeksarray/display-excel-data-in-web-page-using-aspnet-csharp/blob/master/ReadFromExcel/ReadFromExcel/Products.xlsx) - excel file to be shown on web page.
1. [Default.aspx](https://github.com/geeksarray/display-excel-data-in-web-page-using-aspnet-csharp/blob/master/ReadFromExcel/ReadFromExcel/Default.aspx) - Web Page where excel data will be shown. This has DropDownlist to show categories and GridView control to show products.
1. [Default.aspx.cs](https://github.com/geeksarray/display-excel-data-in-web-page-using-aspnet-csharp/blob/master/ReadFromExcel/ReadFromExcel/Default.aspx.cs) - code behind file of web page. This reads content of excel.

## Your web page will have like
![Excel content on ASP.NET Webpage GridView](https://geeksarray.com/images/blog/ProductDetails.png)

For more details visit blog - https://geeksarray/blog/display-excel-data-in-web-page-using-aspnet-csharp
