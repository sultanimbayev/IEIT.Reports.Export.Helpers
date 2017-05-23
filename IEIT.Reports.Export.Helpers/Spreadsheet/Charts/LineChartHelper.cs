using DocumentFormat.OpenXml.Drawing.Charts;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Charts
{
    public static class LineChartHelper
    {
        public static LineChart GetLineChart(this Chart chart)
        {
            return chart.FirstDescendant<LineChart>();
        }

        public static LineChartSeries[] GetSeries(this LineChart lineChart)
        {

            var series = lineChart.Descendants<LineChartSeries>();

            if(series == null || series.Count() == 0)
            {
                return null;
            }
            
            return series.ToArray();
        }

        public static Formula GetValuesFormula(this LineChartSeries series)
        {
            var v = series.FirstDescendant<Values>();
            if(v == null)
            {
                return null;
            }

            var f = v.FirstDescendant<Formula>();
            return f;
        }

        public static Formula GetAxisFormula(this LineChartSeries series)
        {
            var axisData = series.FirstDescendant<CategoryAxisData>();
            if(axisData == null)
            {
                return null;
            }
            var f = axisData.FirstDescendant<Formula>();
            return f;
        }

        public static bool SetValuesFormula(this LineChartSeries series, string newFormulaStr)
        {
            var newFormula = new Formula().From("Drawing\\Charts\\Formula.Empty");
            var newV = new Values().From("Drawing\\Charts\\Values.Empty");
            var newVf = newV.FirstDescendant<Formula>();
            newFormula.Text = newFormulaStr;
            newVf.ReplaceBy(newFormula);

            var oldV = series.FirstDescendant<Values>();
            
            if (oldV == null)
            {
                series.AppendChild(newV);
                return true;
            }

            var oldRef = oldV.FirstDescendant<NumberReference>();
            if (oldRef == null)
            {
                var newRef = newV.FirstDescendant<NumberReference>();
                oldV.PrependChild(newRef);
                return true;
            }
            
            var oldFormula = oldRef.FirstDescendant<Formula>();
            if (oldFormula == null)
            {
                oldRef.PrependChild(newFormula);
                return true;
            }
            
            oldFormula.ReplaceBy(newFormula);
            
            return true;

        }

    }
}
