using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapePropertiesSetFill
    {
        public static void RemoveFill(this xdr.ShapeProperties shapeProperties)
        {
            shapeProperties.RemoveAllChildren<a.SolidFill>();
            shapeProperties.RemoveAllChildren<a.GradientFill>();
            shapeProperties.RemoveAllChildren<a.BlipFill>();
            shapeProperties.RemoveAllChildren<a.GroupFill>();
            var noFillProp = shapeProperties.GetFirstChild<a.NoFill>();
            if(noFillProp == null)
            {
                noFillProp = new a.NoFill();
                shapeProperties.Insert(noFillProp).AfterOneOf(typeof(a.PresetGeometry));
            }
        }
    }
    

}
