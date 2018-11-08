using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DFormatsAddDFormat
    {
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
    }
}
