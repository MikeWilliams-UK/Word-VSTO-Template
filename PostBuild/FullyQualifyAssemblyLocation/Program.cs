using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.Tools.Applications;

namespace FullyQualifyAssemblyLocation
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] parts;

            using (WordprocessingDocument template = WordprocessingDocument.Open(args[0], false))
            {
                var customPropertiesPart = template.CustomFilePropertiesPart;
                var props = template.CustomFilePropertiesPart.Properties;
                var prop = props.FirstChild;

                parts = prop.InnerText.Split('|');
                parts[0] = args[0];
                template.Close();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            // https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-windows-installer?view=vs-2019
            // /assemblyLocation="[INSTALLDIR]ExcelWorkbook.dll"
            // /deploymentManifestLocation="[INSTALLDIR]ExcelWorkbook.vsto"
            // /documentLocation="[INSTALLDIR]ExcelWorkbook.xlsx"
            // /solutionID="Your Solution ID"


            ServerDocument.RemoveCustomization(args[0]);
            //Uri deploymentManifestUri = new Uri(string.Join("|", parts));
            //ServerDocument.AddCustomization(args[0], deploymentManifestUri);
            ServerDocument.AddCustomization(args[0], parts[0], Guid.Parse(parts[1]), new Uri("https://www.chem4word.co.uk"), true, out parts);
            Debugger.Break();
        }
    }
}
