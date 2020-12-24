using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{

    /// <summary>
    /// Extenstion method for setting custom text property
    /// </summary>
    public static class DocumentSetCustomTextProperty
    {

        /// <summary>
        /// Adds or replaces custom text property by name of the custom property
        /// </summary>
        /// <param name="doc">Document to set custom property on</param>
        /// <param name="propertyName">Property name</param>
        /// <param name="propertyValue">Property value</param>
        public static void SetCustomTextProperty(
            this SpreadsheetDocument doc,
            string propertyName,
            string propertyValue)
        {
            // Given a document name, a property name/value, and the property type, 
            // add a custom property to a document. The method returns the original
            // value, if it existed.

            if (string.IsNullOrWhiteSpace(propertyName))
            {
                throw new ArgumentNullException("Custom property name (propertyName) cannot be empty.");
            }

            var newProp = new CustomDocumentProperty();
            newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());


            // Now that you have handled the parameters, start
            // working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            var customProps = doc.CustomFilePropertiesPart;
            if (customProps == null)
            {
                // No custom properties? Add the part, and the
                // collection of properties now.
                customProps = doc.AddCustomFilePropertiesPart();
                customProps.Properties = new Properties();
            }
            var props = customProps.Properties;
            if (props == null)
            {
                props = customProps.Properties = new Properties();
            }

            var prop = props.Select(p => (CustomDocumentProperty)p)
                .Where(p => p.Name.HasValue && p.Name.Value == propertyName)
                .FirstOrDefault();

            // Does the property exist? If so, get the return value, 
            // and then delete the property.
            if (prop != null)
            {
                prop.Remove();
            }

            // Append the new property, and 
            // fix up all the property ID values. 
            // The PropertyId value must start at 2.
            props.AppendChild(newProp);
            int pid = 2;
            foreach (CustomDocumentProperty item in props)
            {
                item.PropertyId = pid++;
            }
            props.Save();
        }
    }
}
