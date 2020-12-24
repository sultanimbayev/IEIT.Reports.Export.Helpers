using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Extenstion methods for getting custom property by property name as string
    /// </summary>
    public static class DocumentGetCustomPropertyAsString
    {
        /// <summary>
        /// Getting custom property by property name as string. Returns null if property not found or doesn't has value.
        /// </summary>
        /// <param name="doc">Document in which the appropriate information will be searched.</param>
        /// <param name="propertyName">Propety name of the custom property.</param>
        /// <returns>Returns null if property not found or doesn't has value. Otherwise, returns value of the custom property.</returns>
        public static string GetCustomPropertyAsString(this SpreadsheetDocument doc, string propertyName)
        {
            if (propertyName == null)
            {
                return null;
            }
            var prop = doc.CustomFilePropertiesPart?.Properties?
                .Select(p => (CustomDocumentProperty)p)
                .FirstOrDefault(p => p.Name.HasValue && p.Name.Value == propertyName);

            return prop?.FirstChild?.InnerText;
        }
    }
}
