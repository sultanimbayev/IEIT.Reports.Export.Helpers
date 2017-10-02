# IEIT.Reports.Export.Helpers

Read [english version](ReadmeEng.md)


Эта библиотека является расширением библиотеки [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/). 
И предназначена для формирования отчетов в формате Microsoft Excel 2007 и выше (.xlsx). Не поддерживает формат Microsoft Excel ниже версии 2007 (файлы формата.xls)
Эта библиотека создана для того, чтобы упростить работу с Excel файлами используя только DocumentFormat.OpenXml.

## Установка с помощью NuGet
```
PM> Install-Package IEIT.Reports.Export.Helpers
```

## Зачем мне это?
С данной библиотекой вы работаете напрямую с драйвером DocumentFormat.OpenXml. Остальные методы являются лишь расширением для существующих классов.
Например, для того чтобы открыть документ:
```C#
var filePath = "myFolder/excelFile.xlsx";
var editable = true;
var excelDoc = SpreadsheetDocument.Open(filePath, editable);
```
Как видите, тут мы используем DocumentFormat.OpenXml, и ничего больше. 

С данным расширениеми операции становятся проще.

Запись в файл:
```C#
var worksheet = excelDoc.GetWorksheet("Лист 1");
worksheet.Write("Привет мир!").To("B2");
excelDoc.SaveAndClose();
```

Получение и присваивание стилей:
```C#
var existingStyleIndex = worksheet.GetCell("A1").StyleIndex;
var cell = worksheet.MakeCell("A4");
cell.StyleIndex = existingStyleIndex;
```

Или можно присвоить стили при записи в ячейку:
```C#
worksheet.Write("Привет мир!").To("B2").WithStyle(existingStyleIndex);
```

Объединение ячеек:
```C#
worksheet.MergeCells("A2:B4");
```