using System;
using System.Collections.Generic;

using System.IO;
using System.Xml;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;

namespace Utilities.Office.Excel
{
    internal class CustomOpenXmlWriter<T> : IDisposable where T : OpenXmlPackaging.OpenXmlPart, OpenXmlPackaging.IFixedContentTypePart
    {
        private XmlWriter Writer { get; set; }
        private IEnumerable<KeyValuePair<string, string>> NamespaceDeclarations { get; set; }
       
        // write to existing part
        public CustomOpenXmlWriter(T openXmlPart)
        {
            // Need to create a fresh part to enable the XmlWriter to write to it if no RootElement exists (i.e. there is no existing document structure preventing the writer)
            IEnumerable<KeyValuePair<string, string>> namespaceDeclarations = null;
            if (openXmlPart != null)
            {
                if (openXmlPart.RootElement != null)
                    openXmlPart = OpenXmlUtilities.RecreatePart<T>(openXmlPart, out namespaceDeclarations);
            }
            NamespaceDeclarations = namespaceDeclarations;

            // This stream is closed when the Writer is closed, so it's OK to use like this
            Stream stream = openXmlPart.GetStream(FileMode.Open, FileAccess.Write);
            Writer = XmlWriter.Create(stream);

            Writer.WriteStartDocument(true);
        }        
        
        public void Close()
        {
            Writer.Close();
        }
        public void Dispose()
        {
            Close();
        }
        public void WriteEndElement()
        {
            Writer.WriteEndElement();
        }
        public void WriteString(string text)
        {
            Writer.WriteString(text);
        }

        // NEW METHODS
        public void WriteOpenXmlElement(OpenXml.OpenXmlElement openXmlElement, bool writeEndElement = false)
        {               
            Writer.WriteStartElement(openXmlElement.LocalName, openXmlElement.NamespaceUri);
            WriteNamespaceDeclarationsAsAttributes();

            if (writeEndElement)
                WriteEndElement();
        }

        public void WriteAttribute(string name, string value)
        {
            Writer.WriteAttributeString(name, value);
        }
        public void WriteAttribute(string name, string value, string prefix)
        {
            Writer.WriteAttributeString(prefix, name, null, value);
        }
        public void WriteAttribute(string name, int value)
        {
            WriteAttribute(name, value.ToString());
        }
        public void WriteAttribute(string name, double value)
        {
            WriteAttribute(name, value.ToString());
        }
        public void WriteAttribute(string name, bool value)
        {
            WriteAttribute(name, (value ? "1" : "0"));
        }
        private void WriteNamespaceDeclarationsAsAttributes()
        {
            if (NamespaceDeclarations != null)
            {
                foreach (KeyValuePair<string, string> namespaceDeclaration in NamespaceDeclarations)
                {
                    WriteAttribute(namespaceDeclaration.Key, namespaceDeclaration.Value, "xmlns");
                }

                NamespaceDeclarations = null;
            }
        }

        public void WriteText(string text)
        {
            Writer.WriteString(text);
        }
        public void WriteText(bool value)
        {
            WriteText((value ? "1" : "0"));
        }


        public static CustomOpenXmlWriter<T> Create(T openXmlPart)
        {
            return new CustomOpenXmlWriter<T>(openXmlPart);
        }
        
    }        
  
}
