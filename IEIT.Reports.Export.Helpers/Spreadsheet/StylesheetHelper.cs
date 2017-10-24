using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetHelper
    {

        /// <summary>
        /// Получить таблицу стилей
        /// </summary>
        /// <param name="document">Рабочий документ</param>
        /// <returns>Таблица стилей указанного документа</returns>
        public static Stylesheet GetStylesheet(this SpreadsheetDocument document)
        {
            return document.WorkbookPart.GetStylesheet();
        }

        /// <summary>
        /// Получить таблицу стилей
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <returns>Таблица стилей указанной рабочей книги</returns>
        public static Stylesheet GetStylesheet(this Workbook workbook)
        {
            return workbook.WorkbookPart.GetStylesheet();
        }

        /// <summary>
        /// Получить таблицу со стилями
        /// </summary>
        /// <param name="wbPart">Рабочая книга документа</param>
        /// <returns>Таблицу содержащяя стили документа</returns>
        public static Stylesheet GetStylesheet(this WorkbookPart wbPart)
        {
            if (wbPart == null) { throw new ArgumentNullException("WorkbookPart must be not null!"); }

            if (wbPart.WorkbookStylesPart == null) { wbPart.AddNewPart<WorkbookStylesPart>(); }
            if (wbPart.WorkbookStylesPart.Stylesheet == null) { wbPart.WorkbookStylesPart.Stylesheet = new Stylesheet(); }
            return wbPart.WorkbookStylesPart.Stylesheet;
        }

        /// <summary>
        /// Добавить формат для условного форматирования ячеек.
        /// Возвращает индекс добавленного формата.
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="format">Новый формат</param>
        /// <returns>Индекс добавленного формата</returns>
        public static uint AddDFormat(this Stylesheet stylesheet, DifferentialFormat format)
        {
            if (stylesheet.DifferentialFormats == null)
            {
                stylesheet.DifferentialFormats = new DifferentialFormats() { Count = 0 };
            }
            return stylesheet.DifferentialFormats.AddDFormat(format);
        }

        /// <summary>
        /// Добавить формат для условного форматирования ячеек.
        /// Возвращает индекс добавленного формата.
        /// </summary>
        /// <param name="formatsList">Объект содержащий форматы данного типа</param>
        /// <param name="format">Новый формат</param>
        /// <returns>Индекс добавленного формата</returns>
        public static uint AddDFormat(this DifferentialFormats formatsList, DifferentialFormat format)
        {
            formatsList.Append(format);
            if (formatsList.Count != null) { formatsList.Count.Value++; }
            return (uint)format.Index();
        }

        #region BordersHelper

        /// <summary>
        /// Получить объект стиля содержащий элементы границ ячеек
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>Объект содержащий элементы границ ячеек</returns>
        internal static Borders GetBorders(this Stylesheet stylesheet)
        {
            if (stylesheet.Borders == null) { stylesheet.Borders = new Borders(new Border()) { Count = 1 }; } // blank border list, if not exists
            return stylesheet.Borders;
        }

        /// <summary>
        /// Получить объект стиля границ ячейки
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="borderIndex">Индекс объекта границ ячеек</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Stylesheet stylesheet, int borderIndex)
        {
            return stylesheet.GetBorders().GetBorder(borderIndex);
        }


        /// <summary>
        /// Получить объект стиля границ ячейки
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="borderIndex">Индекс объекта стиля границ ячеек</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Stylesheet stylesheet, uint borderIndex)
        {
            return stylesheet.GetBorders().GetBorder(borderIndex);
        }


        /// <summary>
        /// Получить объект стиля границ ячейки
        /// </summary>
        /// <param name="borders">Оъект содержащий элементы стлия границ ячеек</param>
        /// <param name="borderIndex">Индекс объекта</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Borders borders, int borderIndex)
        {
            var border = borders.Elements().ElementAt(borderIndex) as Border;
            return border;
        }

        /// <summary>
        /// Получить объект стиля границ ячейки
        /// </summary>
        /// <param name="borders">Оъект содержащий элементы стиля границ ячеек</param>
        /// <param name="borderIndex">Индекс объекта</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Borders borders, uint borderIndex)
        {
            return borders.GetBorder((int)borderIndex);
        }

        /// <summary>
        /// Создать стиль границы ячеек. Возвращает индекс 
        /// созданного стиля.
        /// Не создает обект если такой стиль уже иммется, и
        /// возвращает индекс уже созданного стиля.
        /// </summary>
        /// <param name="borders">Оъект содержащий элементы стиля границ ячеек</param>
        /// <param name="border">Стиль границ ячеек</param>
        /// <returns>
        /// Возвращает индекс созданного стиля, или индекс имееющегося стиля.
        /// </returns>
        public static uint MakeBorder(this Borders borders, Border border)
        {
            var borderIndex = borders.MakeSame(border);
            borders.Count = (uint)borders.Elements().Count();
            return borderIndex;
        }

        /// <summary>
        /// Создать стиль границы ячеек. Возвращает индекс 
        /// созданного стиля.
        /// Не создает обект если такой стиль уже иммется, и
        /// возвращает индекс уже созданного стиля.
        /// </summary>
        /// <param name="stylesheet">Таблица границ ячеек</param>
        /// <param name="border">Стиль границ ячеек</param>
        /// <returns>Возвращает индекс созданного стиля, или индекс имееющегося стиля.</returns>
        public static uint MakeBorder(this Stylesheet stylesheet, Border border)
        {
            return stylesheet.GetBorders().MakeBorder(border);
        }

        #endregion

        #region FontsHelper

        internal static Fonts GetFonts(this Stylesheet stylesheet)
        {
            if(stylesheet.Fonts == null) { stylesheet.Fonts = new Fonts(new Font()) { Count = 1 }; } // blank font list, if not exists
            return stylesheet.Fonts;
        }

        public static Font GetFont(this Stylesheet stylesheet, int fontIndex)
        {
            return stylesheet.GetFonts().GetFont(fontIndex);
        }

        public static Font GetFont(this Stylesheet stylesheet, uint fontIndex)
        {
            return stylesheet.GetFonts().GetFont(fontIndex);
        }

        public static Font GetFont(this Fonts fonts, int fontIndex)
        {
            return fonts.Elements<Font>().ElementAt(fontIndex);
        }

        public static Font GetFont(this Fonts fonts, uint fontIndex)
        {
            return fonts.GetFont((int)fontIndex);
        }

        public static uint MakeFont(this Fonts fonts, Font font)
        {
            var fontIndex = fonts.MakeSame(font);
            fonts.Count = (uint)fonts.Elements().Count();
            return fontIndex;
        }

        public static uint MakeFont(this Stylesheet stylesheet, Font font)
        {
            return stylesheet.GetFonts().MakeFont(font);
        }

        #endregion

        #region NumFormatsHelper

        internal static NumberingFormats GetNumFormats(this Stylesheet stylesheet)
        {
            if(stylesheet.NumberingFormats == null) { stylesheet.NumberingFormats = new NumberingFormats() { Count = 0 }; }
            return stylesheet.NumberingFormats;
        }

        public static NumberingFormat GetNumFormat(this NumberingFormats numFormats, int formatIndex)
        {
            return numFormats.Elements<NumberingFormat>().ElementAt(formatIndex);
        }

        public static NumberingFormat GetNumFormat(this NumberingFormats numFormats, uint formatIndex)
        {
            return numFormats.GetNumFormat((int)formatIndex);
        }

        public static NumberingFormat GetNumFormat(this Stylesheet stylesheet, int formatIndex)
        {
            return stylesheet.GetNumFormats().GetNumFormat(formatIndex);
        }

        public static NumberingFormat GetNumFormat(this Stylesheet stylesheet, uint formatIndex)
        {
            return stylesheet.GetNumFormats().GetNumFormat(formatIndex);
        }

        public static uint MakeNumFormat(this NumberingFormats numFormats, NumberingFormat numFormat)
        {
            var formatIndex = numFormats.MakeSame(numFormat);
            numFormats.Count = (uint)numFormats.Elements().Count();
            return formatIndex;
        }

        public static uint MakeNumFormat(this Stylesheet stylesheet, NumberingFormat numFormat)
        {
            return stylesheet.GetNumFormats().MakeNumFormat(numFormat);
        }

        #endregion

        #region FillsHelper

        internal static Fills GetFills(this Stylesheet stylesheet)
        {
            if(stylesheet.Fills == null) {
                stylesheet.Fills = new Fills() { Count = 2 };
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            }
            return stylesheet.Fills;
        }

        public static Fill GetFill(this Fills fills, int fillIndex)
        {
            return fills.Elements<Fill>().ElementAt(fillIndex);
        }

        public static Fill GetFill(this Fills fills, uint fillIndex)
        {
            return fills.GetFill((int)fillIndex);
        }

        public static Fill GetFill(this Stylesheet stylesheet, int fillIndex)
        {
            return stylesheet.GetFills().GetFill(fillIndex);
        }

        public static Fill GetFill(this Stylesheet stylesheet, uint fillIndex)
        {
            return stylesheet.GetFills().GetFill(fillIndex);
        }

        public static uint MakeFill(this Fills fills, Fill fill)
        {
            var fillIndex = fills.MakeSame(fill);
            fills.Count = (uint)fills.Elements().Count();
            return fillIndex;
        }

        public static uint MakeFill(this Stylesheet stylesheet, Fill fill)
        {
            return stylesheet.GetFills().MakeFill(fill);
        }

        #endregion

        #region CellFormatHelper
        
        internal static CellFormats GetCellFormats(this Stylesheet stylesheet)
        {
            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats(new CellFormat()) { Count = 1 }; // if not exists, then create blank cell format list
                stylesheet.CellFormats.AppendChild(new CellFormat()); // empty one for index 0, seems to be required
            }
            return stylesheet.CellFormats;
        }

        public static CellFormat GetCellFormat(this CellFormats cellFormats, int formatIndex)
        {
            return cellFormats.Elements<CellFormat>().ElementAt(formatIndex);
        }

        public static CellFormat GetCellFormat(this CellFormats cellFormats, uint formatIndex)
        {
            return cellFormats.GetCellFormat((int)formatIndex);
        }

        public static CellFormat GetCellFormat(this Stylesheet stylesheet, int formatIndex)
        {
            return stylesheet.GetCellFormats().GetCellFormat(formatIndex);
        }

        public static CellFormat GetCellFormat(this Stylesheet stylesheet, uint formatIndex)
        {
            return stylesheet.GetCellFormats().GetCellFormat(formatIndex);
        }

        public static uint MakeCellFormat(this CellFormats cellFormats, CellFormat format)
        {
            var formatIndex = cellFormats.MakeSame(format);
            cellFormats.Count = (uint)cellFormats.Elements().Count();
            return formatIndex;
        }

        internal static uint MakeCellFormat(this Stylesheet stylesheet, CellFormat format)
        {
            return stylesheet.GetCellFormats().MakeCellFormat(format);
        }

        #endregion


        public static uint MakeCellStyle(this Stylesheet stylesheet, CellFormat cellFormat)
        {
            return stylesheet.MakeCellFormat(cellFormat);   
        }

        public static MakeStyleIntent MakeCellStyle(this Stylesheet stylesheet)
        {
            return new MakeStyleIntent(stylesheet);
        }

    }
}
