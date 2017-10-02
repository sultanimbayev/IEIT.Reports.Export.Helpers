# IEIT.Reports.Export.Helpers

This is an extension for [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/). 
Supports Microsoft Excel 2007+ (.xlsx) file formats. Doesn't support formats less than Microsoft Excel 2007 (.xls files)
This extension meant to simplify work with Excel files using DocumentFormat.OpenXml.

## Install using NuGet
```
PM> Install-Package IEIT.Reports.Export.Helpers
```

## Why do I need this?
This library gives you extensions over DocumentFormat.OpenXml classes.
For example, to open an Excel file you type:
```C#
var filePath = "myFolder/excelFile.xlsx";
var editable = true;
var excelDoc = SpreadsheetDocument.Open(filePath, editable);
```
We just using DocumentFormat.OpenXml syntax to open files. 

This extension simplifies operations with excel files.
Writing to file:
```C#
var worksheet = excelDoc.GetWorksheet("List 1");
worksheet.Write("Hello world!").To("B2");
excelDoc.SaveAndClose();
```


Fetching and assigning styles:
```C#
var existingStyleIndex = worksheet.GetCell("A1").StyleIndex;
var cell = worksheet.MakeCell("A4");
cell.StyleIndex = existingStyleIndex;
```


You can put style into cell when writing:
```C#
worksheet.Write("Привет мир!").To("B2").WithStyle(existingStyleIndex);
```


Merging cells:
```C#
worksheet.MergeCells("A2:B4");
```