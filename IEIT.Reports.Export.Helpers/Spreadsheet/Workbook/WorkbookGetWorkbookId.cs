using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorkbookGetWorkbookId
    {
        public const string WORKBOOK_ID_ATTRIBUTE_NAME = "workbookId";
        public static string GetWorkbookId(this Workbook workbook)
        {
            var workbookIdAttribute = workbook.ExtendedAttributes
                .FirstOrDefault(attr => attr.LocalName == WORKBOOK_ID_ATTRIBUTE_NAME 
                    && string.IsNullOrEmpty(attr.NamespaceUri));

            if(default(OpenXmlAttribute) == workbookIdAttribute)
            {
                var guid = Guid.NewGuid();
                workbookIdAttribute = new OpenXmlAttribute()
                {
                    LocalName = WORKBOOK_ID_ATTRIBUTE_NAME,
                    Value = guid.ToString()
                };
                workbook.SetAttribute(workbookIdAttribute);
            }
            return workbookIdAttribute.Value;
        }
    }
}
