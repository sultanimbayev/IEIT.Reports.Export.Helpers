using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using sysDr = System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapePropsSetSolidFill
    {
        public static void SetSolidFill(this xdr.ShapeProperties shapeProperties, sysDr.Color fillColor, float alpha = 1f)
        {
            shapeProperties.RemoveAllChildren<a.NoFill>();
            shapeProperties.RemoveAllChildren<a.GradientFill>();
            shapeProperties.RemoveAllChildren<a.BlipFill>();
            shapeProperties.RemoveAllChildren<a.GroupFill>();
            shapeProperties.RemoveAllChildren<a.SolidFill>();
            var solidFill = new a.SolidFill();
            shapeProperties.Insert(solidFill).AfterOneOf(typeof(a.PresetGeometry));
            var fillColorModel = new a.RgbColorModelHex();
            fillColorModel.Val = fillColor.ToHex();
            solidFill.Append(fillColorModel);
            var colorAlpha = new a.Alpha() { Val = (int)(alpha * 100000) }; // FillAlpha - def = val / 1000
            fillColorModel.Append(colorAlpha);
        }
    }
}
