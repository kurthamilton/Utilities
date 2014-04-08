using System;
using System.Collections.ObjectModel;

using System.IO;
using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    
    internal class CustomOpenXmlAttribute
    {
        private OpenXml.OpenXmlAttribute Attribute { get; set; }

        public string LocalName { get { return Attribute.LocalName; } }
        public string Value { get { return Attribute.Value; } }

        public CustomOpenXmlAttribute(OpenXml.OpenXmlAttribute attribute)
        {
            Attribute = attribute;
        }

        public int GetIntValue()
        {
            return int.Parse(Value);
        }
        public bool GetBoolValue()
        {
            return (Value == "1");
        }
        public double GetDoubleValue()
        {
            return double.Parse(Value);
        }
    }        
  
}
