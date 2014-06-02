using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;

using OpenXml = DocumentFormat.OpenXml;

namespace Utilities.Office.Excel
{    
    internal class CustomOpenXmlAttributeCollection : IEnumerable<CustomOpenXmlAttribute>
    {
        private List<CustomOpenXmlAttribute> Attributes { get; set; }

        public CustomOpenXmlAttributeCollection(OpenXml.OpenXmlReader reader)
        {
            Attributes = new List<CustomOpenXmlAttribute>();

            foreach (OpenXml.OpenXmlAttribute attribute in reader.Attributes)
            {
                Attributes.Add(new CustomOpenXmlAttribute(attribute));
            }
        }

        public CustomOpenXmlAttribute this[string localName]
        {
            get
            {
                return GetAttribute(localName);
            }
        }

        public CustomOpenXmlAttribute GetAttribute(string localName)
        {
            foreach (CustomOpenXmlAttribute attribute in Attributes)
            {
                if (attribute.LocalName == localName)
                    return attribute;
            }
            return null;
        }

        public bool TryGetBoolAttributeValue(string localName)
        {
            CustomOpenXmlAttribute attribute = GetAttribute(localName);
            if (attribute != null)
                return attribute.GetBoolValue();
            return false;
        }

        // Implement IEnumerable
        public IEnumerator<CustomOpenXmlAttribute> GetEnumerator()
        {
            return Attributes.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }        
  
}
