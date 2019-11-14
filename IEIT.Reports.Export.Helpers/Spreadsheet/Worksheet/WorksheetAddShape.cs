using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;
using x =DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetAddShape
    {
        public static xdr.Shape AddShape(this x.Worksheet ws, a.ShapeTypeValues shapeType, string name = null)
        {
            var worksheetDrawing = ws.GetWorksheetDrawing();
            return worksheetDrawing.AddShape(shapeType, name); ;
        }
    }
}
