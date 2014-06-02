using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal static class WorkbookCleaner
    {

        public static bool CleanWorkbook(string filePath)
        {
            using (Package excelPackage = TryOpenPackage(filePath))
            {
                CleanWorkbookRelsDocument(excelPackage);                
            }
            return false;                
        }        


        private static void CleanWorkbookRelsDocument(Package excelPackage)
        {
            PackagePart workbookRelsPart = GetPackagePart(excelPackage, @"xl\_rels\workbook.xml.rels");
            XmlDocument workbookRelsDocument = GetXmlDocumentFromPackagePart(workbookRelsPart);

            foreach (XmlNode workbookRelNode in workbookRelsDocument.DocumentElement.ChildNodes)
            {
                XmlAttribute typeAttribute = workbookRelNode.Attributes["Type"];

                if (typeAttribute.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                {
                    string worksheetFilePath = workbookRelNode.Attributes["Target"].Value;
                    string sheetName = new FileInfo(worksheetFilePath).Name;

                    CleanWorksheetRelsDocument(excelPackage, sheetName);
                }
            }
        }

        private static void CleanWorksheetRelsDocument(Package excelPackage, string sheetName)
        {
            string worksheetRelsFilePath = @"xl\worksheets\_rels\" + sheetName + ".rels";
            PackagePart worksheetRelsPart = GetPackagePart(excelPackage, worksheetRelsFilePath);

            if (worksheetRelsPart != null)
            {
                XmlDocument worksheetRelsDocument = GetXmlDocumentFromPackagePart(worksheetRelsPart);

                List<XmlNode> nodesToRemove = new List<XmlNode>();

                foreach (XmlNode worksheetRelsNode in worksheetRelsDocument.DocumentElement.ChildNodes)
                {                    
                    if (worksheetRelsNode.Attributes["TargetMode"] != null)
                    {
                        // remove relationships to external files
                        if (worksheetRelsNode.Attributes["TargetMode"].Value == "External")
                        {
                            nodesToRemove.Add(worksheetRelsNode);
                        }
                    }
                }

                if (nodesToRemove.Count > 0)
                {
                    RemoveNodes(nodesToRemove);
                    SaveXmlDocumentToPackagePart(worksheetRelsPart, worksheetRelsDocument);
                }
            }
        }

        private static Package TryOpenPackage(string filePath)
        {
            if (File.Exists(filePath))
            {
                DateTime startTime = DateTime.Now;

                do
                {
                    try
                    {
                        Package package = Package.Open(filePath, FileMode.Open);
                        return package;
                    }
                    catch (IOException ex)
                    {
                        string message = ex.Message;
                    }
                }
                while ((DateTime.Now - startTime) < TimeSpan.FromSeconds(10));
            }

            return null;
        }

        private static PackagePart GetPackagePart(Package package, string partFilePath)
        {
            Uri packagePartUri = GetPackagePartUri(partFilePath);

            if (package.PartExists(packagePartUri))
                return package.GetPart(packagePartUri);
            else
                return null;
        }

        private static Uri GetPackagePartUri(string partFilePath)
        {
            return PackUriHelper.CreatePartUri(new Uri(partFilePath, UriKind.Relative));
        }        

        private static XmlDocument GetXmlDocumentFromPackagePart(PackagePart packagePart)
        {
            XmlDocument xmlDocument = null;
            if (packagePart != null)
            {
                using (Stream packagePartStream = packagePart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    using (StreamReader streamReader = new StreamReader(packagePartStream)) 
                    {
                        xmlDocument = new XmlDocument();
                        xmlDocument.LoadXml(streamReader.ReadToEnd());                        
                    }
                }
            }

            return xmlDocument;
        }

        private static void RemoveNodes(List<XmlNode> nodes)
        {
            foreach (XmlNode node in nodes)
            {
                node.ParentNode.RemoveChild(node);
            }
        }

        private static void SaveXmlDocumentToPackagePart(PackagePart packagePart, XmlDocument xmlDocument)
        {
            using (Stream packagePartStream = packagePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                xmlDocument.Save(packagePartStream);
            }
        }
    }
}
