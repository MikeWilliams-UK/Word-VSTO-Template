using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace FullyQualifyAssemblyLocation
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordprocessingDocument template = WordprocessingDocument.Open(args[0], true))
            {
                var customPropertiesPart = template.CustomFilePropertiesPart;
                var props = template.CustomFilePropertiesPart.Properties;
                var prop = props.FirstChild;

                var parts = prop.InnerText.Split('|');
                parts[0] = args[0];
                //prop.InnerText = string.Join('|', parts);
                
                Debugger.Break();
            }
        }
    }
}
