﻿using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace FullyQualifyAssemblyLocation
{
    public static class Program
    {
        private static void Main(string[] args)
        {
            string[] parts = null;

            using (WordprocessingDocument template = WordprocessingDocument.Open(args[0], true))
            {
                if (template.CustomFilePropertiesPart != null)
                {
                    var customProperties = template.CustomFilePropertiesPart.Properties;

                    foreach (var customProperty in customProperties)
                    {
                        if (customProperty.InnerText.Contains("|vstolocal"))
                        {
                            parts = customProperty.InnerText.Split('|');
                            parts[0] = args[0].Replace(".dotx", ".vsto");
                        }
                    }

                    if (parts != null)
                    {
                        customProperties.RemoveAllChildren();
                        customProperties.Append(NewProperty(2, "_AssemblyLocation", string.Join("|", parts)));
                        customProperties.Append(NewProperty(3, "_AssemblyName", "4E3C66D5-58D4-491E-A7D4-64AF99AF6E8B"));

                        template.Save();
                        template.Close();
                    }
                }
            }
        }

        private static OpenXmlElement[] NewProperty(int propertyId, string propertyName, string propertyValue)
        {
            var property = new CustomDocumentProperty
            {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = propertyId,
                Name = propertyName
            };

            var vTlpwstr1 = new Vt.VTLPWSTR
            {
                Text = propertyValue
            };
            property.Append(vTlpwstr1.ToArray());

            return property.ToArray();
        }
    }
}