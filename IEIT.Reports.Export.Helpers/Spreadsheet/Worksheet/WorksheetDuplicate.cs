using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DuplicateWorksheet
    {
        /// <summary>
        /// Копирует лист
        /// </summary>
        /// <param name="ws">Исходный лист</param>
        /// <param name="newSheetName">Имя нового листа</param>
        /// <param name="docType">Тип исходного листа SpreadsheetDocumentType</param>
        public static void Duplicate(this Worksheet ws, string newSheetName, SpreadsheetDocumentType docType = SpreadsheetDocumentType.Workbook)
        {
            var sourceSheetPart = ws.WorksheetPart;
            SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), docType);
            WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();
            WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart(sourceSheetPart);

            var WbPart = ws.GetWorkbookPart();
            //Add cloned sheet and all associated parts to workbook
            WorksheetPart clonedSheet = WbPart.AddPart<WorksheetPart>(tempWorksheetPart);
            //Table definition parts are somewhat special and need unique ids...so let's make an id based on count
            int numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();
            //Clean up table definition parts (tables need unique ids)
            if (numTableDefParts != 0)
            {
                //Every table needs a unique id and name
                foreach (TableDefinitionPart tableDefPart in clonedSheet.TableDefinitionParts)
                {
                    numTableDefParts++;
                    tableDefPart.Table.Id = (uint)numTableDefParts;
                    tableDefPart.Table.DisplayName = "CopiedTable" + numTableDefParts;
                    tableDefPart.Table.Name = "CopiedTable" + numTableDefParts;
                    tableDefPart.Table.Save();
                }
            }

            //There can only be one sheet that has focus
            SheetViews views = clonedSheet.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                clonedSheet.Worksheet.Save();
            }

            //Add new sheet to main workbook part
            Sheets sheets = WbPart.Workbook.GetFirstChild<Sheets>();
            Sheet copiedSheet = new Sheet
            {
                Name = newSheetName,
                Id = WbPart.GetIdOfPart(clonedSheet),
                SheetId = (uint)sheets.ChildElements.Count + 1
            };
            sheets.Append(copiedSheet);
            //Save Changes
            WbPart.Workbook.Save();
        }
    }
}
