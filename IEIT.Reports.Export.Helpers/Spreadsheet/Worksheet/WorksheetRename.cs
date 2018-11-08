using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetRename
    {
        /// <summary>
        /// Переименовать лист
        /// </summary>
        /// <param name="worksheet">Лист который ты хочешь переименовать</param>
        /// <param name="newName">Новое название листа</param>
        /// <param name="updateReferences">Заменить все ссылки к данному листу?</param>
        /// <returns>true при удачном переименовывании, false в обратном случае</returns>
        public static bool Rename(this Worksheet worksheet, string newName, bool updateReferences = false)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return false; }

            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return false; }

            if (updateReferences)
            {
                var pattern = $"'?{sheet.Name}'?!";
                var replacement = $"'{newName}'!";

                var wsParts = wbPart.GetPartsOfType<WorksheetPart>();
                foreach (var wsPart in wsParts)
                {
                    //TODO: Улучшить обновление ссылок на лист
                    wsPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    foreach (var chartPart in wsPart.DrawingsPart.ChartParts)
                    {
                        chartPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    }
                }
            }

            return (sheet.Name = newName).Equals(newName);
        }
    }
}
