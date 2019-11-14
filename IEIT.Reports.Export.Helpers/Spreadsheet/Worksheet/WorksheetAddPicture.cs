using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;
using x = DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetAddPicture
    {
        public static xdr.Picture AddPicture(this x.Worksheet ws, string imagePath)
        {
            var worksheetDrawing = ws.GetWorksheetDrawing();
            return worksheetDrawing.AddPicture(imagePath);
        }
    }
}
