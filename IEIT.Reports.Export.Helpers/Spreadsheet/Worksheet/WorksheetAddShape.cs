using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using x =DocumentFormat.OpenXml.Spreadsheet;
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
        public static xdr.Shape AddShape(this x.Worksheet ws, string name, a.ShapeTypeValues shapeType)
        {

            var drawingsPart = ws.WorksheetPart.GetPartsOfType<DrawingsPart>().FirstOrDefault();
            if(drawingsPart == null)
            {
                var count = ws.WorksheetPart.Parts.Count();
                drawingsPart = ws.WorksheetPart.AddNewPart<DrawingsPart>("rId" + (count + 1));

            }
            var drawingsPartId = ws.WorksheetPart.GetIdOfPart(drawingsPart);
            var worksheetDrawing = drawingsPart.WorksheetDrawing;
            if(worksheetDrawing == null)
            {
                worksheetDrawing = new xdr.WorksheetDrawing();
                worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                drawingsPart.WorksheetDrawing = worksheetDrawing;
            }
            uint? lastDrawingId = worksheetDrawing.Descendants<xdr.NonVisualDrawingProperties>().Select(p => p.Id?.Value).Max();
            

            var twoCellAnchor = new xdr.TwoCellAnchor().InitDefault();
            worksheetDrawing.Append(twoCellAnchor);
                        
            var shape = new xdr.Shape() { Macro = "", TextLink = "" };
            twoCellAnchor.Append(shape);
            shape.NonVisualShapeProperties = new xdr.NonVisualShapeProperties();

            var drawingProps = new xdr.NonVisualDrawingProperties();
            shape.NonVisualShapeProperties.NonVisualDrawingProperties = drawingProps;
            drawingProps.Id = (lastDrawingId ?? 0) + 1; //ID - auto
            
            drawingProps.Name = name ?? Guid.NewGuid().ToString();
            
            shape.NonVisualShapeProperties.NonVisualShapeDrawingProperties = new xdr.NonVisualShapeDrawingProperties();

            shape.ShapeProperties = new xdr.ShapeProperties().Init(shapeType); //ShapeProps - user def
            
            var shapeStyle = new xdr.ShapeStyle().InitDefault();
            shape.Append(shapeStyle);

            var txtBody = new xdr.TextBody().InitDefault(); //Text - user def
            shape.Append(txtBody);
            
            twoCellAnchor.Append(new xdr.ClientData());

            var drawing = ws.GetFirstChild<x.Drawing>();
            if(drawing == null)
            {
                drawing = new x.Drawing()
                {
                    Id = drawingsPartId
                };
                ws.Append(drawing);
            }
            var wbPart = ws.WorksheetPart.ParentPartOfType<WorkbookPart>();
            wbPart.Workbook.Save();

            return shape;
        }
    }
}
