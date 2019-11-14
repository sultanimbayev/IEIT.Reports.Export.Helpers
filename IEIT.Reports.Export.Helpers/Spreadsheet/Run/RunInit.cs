using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;


namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RunInit
    {
        public static a.Run Init(this a.Run run, string text, Font font, Color? fontColor = null)
        {
            run.Text = new a.Text(text);
            var runProps = new a.RunProperties();
            run.Append(runProps);
            runProps.SetFont(font ?? new Font("Calibri", 11), fontColor);
            return run;
        }
    }
}
