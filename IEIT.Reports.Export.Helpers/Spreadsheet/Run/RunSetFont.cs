using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RunSetFont
    {
        public static a.Run SetFont(this a.Run run, Font font, Color? fontColor = null)
        {
            var runProps = run.GetFirstChild<a.RunProperties>();
            if(runProps == null)
            {
                runProps = new a.RunProperties();
                run.PrependChild(runProps);
            }
            runProps.SetFont(font, fontColor);
            return run;
        }
    }
}
