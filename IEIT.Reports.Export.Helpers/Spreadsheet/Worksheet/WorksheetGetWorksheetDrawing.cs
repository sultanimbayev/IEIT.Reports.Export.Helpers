using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;
using x = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetWorksheetDrawing
    {
        public static xdr.WorksheetDrawing GetWorksheetDrawing(this x.Worksheet ws)
        {
            var drawingsPart = ws.WorksheetPart.GetPartsOfType<DrawingsPart>().FirstOrDefault();
            if (drawingsPart == null)
            {
                var count = ws.WorksheetPart.Parts.Count();
                drawingsPart = ws.WorksheetPart.AddNewPart<DrawingsPart>("rId" + (count + 1));
            }
            var drawingsPartId = ws.WorksheetPart.GetIdOfPart(drawingsPart);
            var worksheetDrawing = drawingsPart.WorksheetDrawing;
            if (worksheetDrawing == null)
            {
                worksheetDrawing = new xdr.WorksheetDrawing();
                worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                drawingsPart.WorksheetDrawing = worksheetDrawing;
            }
            var drawing = ws.GetFirstChild<x.Drawing>();
            if (drawing == null)
            {
                drawing = new x.Drawing()
                {
                    Id = drawingsPartId
                };
                ws.Append(drawing);
            }

            var wbPart = ws.WorksheetPart.ParentPartOfType<WorkbookPart>();
            wbPart.Workbook.Save();

            return worksheetDrawing;
        }
    }
}
