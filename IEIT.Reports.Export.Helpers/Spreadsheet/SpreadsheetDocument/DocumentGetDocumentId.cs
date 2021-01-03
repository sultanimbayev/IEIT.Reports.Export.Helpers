using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Extension methods for generating and retrieveing id of a document
    /// </summary>
    public static class DocumentGetDocumentId
    {
        /// <summary>
        /// Get property name for retrieving a document id
        /// </summary>
        /// <returns>property name for retrieving a document id</returns>
        public static string GetDocumentIdPropertyName()
        {
            var ns = typeof(DocumentGetDocumentId).Namespace;
            return ns + ".DocumentId";
        }

        /// <summary>
        /// Get document id. Generates new document id if not exists.
        /// </summary>
        /// <param name="doc">A document of which to get the id.</param>
        /// <returns></returns>
        public static string GetDocumentId(this SpreadsheetDocument doc)
        {
            var documentIdPropertyName = GetDocumentIdPropertyName();
            var documentId = doc.GetCustomPropertyAsString(documentIdPropertyName);
            if (!string.IsNullOrWhiteSpace(documentId))
            {
                return documentId;
            }
            var newDocumentId = Guid.NewGuid().ToString();
            doc.SetCustomTextProperty(documentIdPropertyName, newDocumentId);
            return newDocumentId;
        }
    }
}
