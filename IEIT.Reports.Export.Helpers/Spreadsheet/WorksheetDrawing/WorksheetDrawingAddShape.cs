using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetDrawingAddShape
    {
        public static xdr.Shape AddShape(this xdr.WorksheetDrawing worksheetDrawing, a.ShapeTypeValues shapeType, string name = null)
        {
            var twoCellAnchor = new xdr.TwoCellAnchor().InitDefault();
            worksheetDrawing.Append(twoCellAnchor);

            var shape = new xdr.Shape() { Macro = "", TextLink = "" };
            twoCellAnchor.Append(shape);
            shape.NonVisualShapeProperties = new xdr.NonVisualShapeProperties();

            var drawingProps = new xdr.NonVisualDrawingProperties();
            shape.NonVisualShapeProperties.NonVisualDrawingProperties = drawingProps;

            uint? lastDrawingId = worksheetDrawing.Descendants<xdr.NonVisualDrawingProperties>().Select(p => p.Id?.Value).Max();
            drawingProps.Id = (lastDrawingId ?? 0) + 1; //ID - auto

            drawingProps.Name = name ?? Guid.NewGuid().ToString();

            shape.NonVisualShapeProperties.NonVisualShapeDrawingProperties = new xdr.NonVisualShapeDrawingProperties();

            shape.ShapeProperties = new xdr.ShapeProperties().Init(shapeType); //ShapeProps - user def

            var shapeStyle = new xdr.ShapeStyle().InitDefault();
            shape.Append(shapeStyle);

            var txtBody = new xdr.TextBody().InitDefault(); //Text - user def
            shape.Append(txtBody);

            twoCellAnchor.Append(new xdr.ClientData());

            return shape;
        }
    }
}
