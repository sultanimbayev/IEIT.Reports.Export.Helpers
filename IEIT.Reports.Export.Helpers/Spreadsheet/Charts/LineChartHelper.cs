using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.Generic;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Charts
{
    public static class LineChartHelper
    {

        /// <summary>
        /// Получить вложенный объект линейного графика из объекта диаграммы
        /// </summary>
        /// <param name="chart">Объект содержащий график</param>
        /// <returns>Линейный график, если в данном объекте диаграммы он есть. В обратном случае возвращается <code>null</code></returns>
        public static LineChart AsLineChart(this Chart chart)
        {
            return chart.FirstDescendant<LineChart>();
        }

        /// <summary>
        /// Получить все ряды из линейного графика
        /// </summary>
        /// <param name="lineChart">Линейный график</param>
        /// <returns>Ряды из данного линейного графика, или <code>null</code> если их нет</returns>
        public static IEnumerable<LineChartSeries> Series(this LineChart lineChart)
        {
            var series = lineChart.Descendants<LineChartSeries>();
            if(series == null || series.Count() == 0) { return null; }
            return series;
        }


        /// <summary>
        /// Получить формулу значении
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <returns>Формула значении ряда линейного графика</returns>
        public static Formula Values(this LineChartSeries series)
        {
            var v = series.FirstDescendant<Values>();
            if(v == null) { return null; }
            var f = v.FirstDescendant<Formula>();
            return f;
        }

        /// <summary>
        /// Задать формулу значении ряда
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <param name="newFormula">Новая формула значении</param>
        /// <returns>true - если формула успешно задана, false - в обратном случае</returns>
        public static bool Values(this LineChartSeries series, Formula newFormula)
        {
            var newV = new Values() { NumberReference = new NumberReference() { Formula = newFormula } };
            var oldV = series.FirstDescendant<Values>();
            var newElem = oldV.ReplaceBy(newV);
            return newElem.SameAs(newV);
        }

        /// <summary>
        /// Задать формулу ряда
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <param name="newFormulaStr">Новая формула</param>
        /// <returns>true - если формула успешно задана, false - в обратном случае</returns>
        public static bool Values(this LineChartSeries series, string newFormulaStr)
        {
            var newFormula = new Formula() { Text = newFormulaStr };
            return series.Values(newFormula);
        }

        /// <summary>
        /// Получить формулу значении (горизонтальной) оси
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <returns>Формула значении оси ряда линейного графика</returns>
        public static Formula AxisValues(this LineChartSeries series)
        {
            var axisData = series.FirstDescendant<CategoryAxisData>();
            if(axisData == null) { return null; }
            var f = axisData.FirstDescendant<Formula>();
            return f;
        }

        /// <summary>
        /// Задать формулу значении (горизонтальной) оси
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <param name="newFormula">Новая формула значении</param>
        /// <returns>true - если формула успешно задана, false - в обратном случае</returns>
        public static bool AxisValues(this LineChartSeries series, Formula newFormula)
        {
            var oldValues = series.FirstDescendant<CategoryAxisData>();
            var newV = new CategoryAxisData() { NumberReference = new NumberReference() { Formula = newFormula } };
            var newElem = oldValues.ReplaceBy(newV);
            return newElem.SameAs(newV);
        }

        /// <summary>
        /// Задать формулу значении (горизонтальной) оси
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <param name="newFormulaStr">Новая формула значении в виде строки</param>
        /// <returns>true - если формула успешно задана, false - в обратном случае</returns>
        public static bool AxisValues(this LineChartSeries series, string newFormulaStr)
        {
            var newFormula = new Formula() { Text = newFormulaStr };
            return series.AxisValues(newFormula);
        }

        /// <summary>
        /// Получить ссылку на ячейки составляющее названия ряда или номер ряда
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <returns>Формулу названия или номер ряда</returns>
        public static string Label(this LineChartSeries series)
        {
            var seriesText = series.SeriesText;
            if(seriesText == null || seriesText.FirstChild == null) { return null; }
            var child = seriesText.FirstChild;
            if(child is StringReference)
            {
                var strRef = child as StringReference;
                return strRef?.Formula?.Text;
            }
            if(child is NumericValue)
            {
                var numVal = child as NumericValue;
                return numVal?.Text;
            }
            return null;
        }


        /// <summary>
        /// Задать название ряда ссылкой на ячейки
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <param name="newFormula">Формула-ссылка названия ряда</param>
        /// <returns>true - если название успешно задано, false - в обратном случае</returns>
        public static bool Label(this LineChartSeries series, Formula newFormula)
        {
            var seriesText = new SeriesText() { StringReference = new StringReference() { Formula = newFormula } };
            series.SeriesText = seriesText;
            return true;
        }

        /// <summary>
        /// Задать название ряда ссылкой на ячейки
        /// </summary>
        /// <param name="series">Ряд значении линейного графика</param>
        /// <param name="newFormulaStr">Формула-ссылка названия ряда</param>
        /// <returns>true - если название успешно задано, false - в обратном случае</returns>
        public static bool Label(this LineChartSeries series, string newFormulaStr)
        {
            var newFormula = new Formula() { Text = newFormulaStr };
            return series.Label(newFormula);
        }

    }
}
