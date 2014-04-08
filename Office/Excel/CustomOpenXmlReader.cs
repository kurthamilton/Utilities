using System;
using System.Collections.Generic;

using System.IO;
using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    
    internal class CustomOpenXmlReader : IDisposable
    {
        private OpenXml.OpenXmlReader Reader { get; set; }

        private CustomOpenXmlAttributeCollection _attributes = null;
        public CustomOpenXmlAttributeCollection Attributes { get { if (_attributes == null) _attributes = new CustomOpenXmlAttributeCollection(Reader); return _attributes; } }
        public int Depth { get { return Reader.Depth; } }
        public Type ElementType { get { return Reader.ElementType; } }
        public string Encoding { get { return Reader.Encoding;  } }
        public bool EOF { get { return Reader.EOF; } }
        public bool IsEndElement { get { return Reader.IsEndElement; } }
        public bool IsStartElement { get { return Reader.IsStartElement; } }
        public string LocalName { get { return Reader.LocalName; } }
        public string NamespaceUri { get { return Reader.NamespaceUri; } }
        public string Prefix { get { return Reader.Prefix; } }

        public OpenXml.OpenXmlElement OpenXmlElement { get; private set; }
        
        private CustomOpenXmlReader(OpenXml.OpenXmlElement openXmlElement)
        {
            OpenXmlElement = openXmlElement;
            Reader = OpenXml.OpenXmlReader.Create(openXmlElement);            
        }
              
        public void Close()
        {
            Reader.Close();
        }
        public void Dispose()
        {
            Reader.Dispose();
        }
        public string GetText()
        {
            return Reader.GetText();
        }
        public OpenXml.OpenXmlElement LoadCurrentElement()
        {
            return Reader.LoadCurrentElement();
        }
        public bool Read()
        {
            _attributes = null;
            return Reader.Read();
        }
        public bool ReadFirstChild()
        {
            _attributes = null;
            return Reader.ReadFirstChild();
        }

        public static CustomOpenXmlReader Create(OpenXml.OpenXmlElement openXmlElement)
        {
            return new CustomOpenXmlReader(openXmlElement);
        }

        // New methods
        public bool IsEndElementOfType<T>() where T : OpenXml.OpenXmlElement
        {
            return (IsEndElement && ElementType == typeof(T));
        }
        public bool IsStartElementOfType<T>() where T : OpenXml.OpenXmlElement
        {
            return (IsStartElement && ElementType == typeof(T));
        }
        public bool ReadToEndElement<T>() where T : OpenXml.OpenXmlElement
        {
            return ReadToEndElement(typeof(T));
        }
        public bool ReadToEndElement(Type elementType)
        {
            bool read = Read();
            if (read)
            {
                if (IsEndElement)
                    read = (ElementType != elementType);
            }
            return read;
        }        

        public static bool GetBoolValue(string value)
        {
            if (value == "1")
                return true;
            return false;
        }
    }        
  
}
