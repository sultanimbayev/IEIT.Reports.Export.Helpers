using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class Fabric
    {
        public static ConditionalFormattingRule MakeFormattingRule(string expression, int priority = 1)
        {
            return new ConditionalFormattingRule(new Formula(expression)) { Type = ConditionalFormatValues.Expression, Priority = priority };
        }
    }
}
