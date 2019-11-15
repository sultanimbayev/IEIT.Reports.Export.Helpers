using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetDrawingAddPicture
    {
        public static xdr.Picture AddPicture(this xdr.WorksheetDrawing worksheetDrawing, string imagePath, string name = null)
        {
            var partsCount = worksheetDrawing.DrawingsPart.Parts.Count();
            var extenstion = Path.GetExtension(imagePath);
            var newPartId = "rId" + (partsCount + 1);
            var imagePart = worksheetDrawing.DrawingsPart.AddNewPart<ImagePart>("image/" + extenstion, newPartId);
            using (var stream1 = File.OpenRead(imagePath))
            {
                imagePart.FeedData(stream1);
            }
            var pict = AddPictureByEmbededPictureRefence(worksheetDrawing, newPartId, name);
            return pict;
        }

        public static xdr.Picture AddPictureByEmbededPictureRefence(this xdr.WorksheetDrawing worksheetDrawing, string refId, string name = null)
        {
            var twoCellAnchor = new xdr.TwoCellAnchor().InitDefault();
            twoCellAnchor.EditAs = xdr.EditAsValues.OneCell;
            worksheetDrawing.Append(twoCellAnchor);

            var pict = new xdr.Picture();
            twoCellAnchor.Append(pict);

            var nonVisualPictureProperties = new xdr.NonVisualPictureProperties();
            pict.Append(nonVisualPictureProperties);

            uint? lastDrawingId = worksheetDrawing.Descendants<xdr.NonVisualDrawingProperties>().Select(p => p.Id?.Value).Max();
            var nonVisualDrawingProperties = new xdr.NonVisualDrawingProperties()
                {
                    Id = (lastDrawingId ?? 0) + 1,
                    Name = name ?? Guid.NewGuid().ToString()
                };
            nonVisualPictureProperties.Append(nonVisualDrawingProperties);

            var nonVisualPictureDrawingProperties = new xdr.NonVisualPictureDrawingProperties();
            nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);

            var pictureLocks = new a.PictureLocks() { NoChangeAspect = true };
            nonVisualPictureDrawingProperties.Append(pictureLocks);

            var blipFill = new xdr.BlipFill();
            blipFill.Blip = new a.Blip() { Embed = refId };
            pict.Append(blipFill);

            var stretch = new a.Stretch();
            stretch.FillRectangle = new a.FillRectangle();
            blipFill.Append(stretch);

            var shapeProps = new xdr.ShapeProperties().InitPictureDefault();
            pict.Append(shapeProps);

            twoCellAnchor.Append(new xdr.ClientData());

            return pict;

        }
    }
}
