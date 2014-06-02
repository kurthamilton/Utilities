using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    
    public class DefinedName : BaseExcel, IEquatable<DefinedName>
    {
        internal const int DefaultDefinedNameId = 0;

        internal Workbook Workbook { get; private set; }
        
        internal int NameId { get; private set; }

        public string Name { get; private set; }                
        public List<string> Addresses { get; private set; }

        public Worksheet Worksheet { get; private set; }
        public DefinedNameScope Scope { get; private set; }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        private DefinedName(Workbook workbook)
        {
            Workbook = workbook;
        }

        internal DefinedName(Workbook workbook, int nameId, string name, string addressString)
            : this(workbook, nameId, name, addressString, DefinedNameScope.Workbook)
        {            
        }

        internal DefinedName(Workbook workbook, int nameId, string name, string addressString, DefinedNameScope scope)
            : this (workbook)
        {
            NameId = nameId;
            Name = name;
            Scope = scope;

            SetAddresses(addressString);
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public void Delete()
        {
            Workbook.DefinedNames.Delete(NameId);
        }

        // implement IEquatable
        public bool Equals(DefinedName other)
        {
            return (string.Compare(other.Name, Name, StringComparison.InvariantCultureIgnoreCase) == 0 &&
                other.Scope == Scope);
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal void SetAddresses(string addressString)
        {
            Worksheet worksheet;
            List<string> addresses;
            GetAddressElementsFromAddressString(Workbook, addressString, out worksheet, out addresses);

            Worksheet = worksheet;
            Addresses = addresses;
        }

        internal string GetAddressString(bool useLocal)
        {            
            if (useLocal)
            {
                return string.Join(",", Addresses.ToArray());
            }
            else
            {
                List<string> addresses = new List<string>();

                foreach (string address in Addresses)
                {
                    addresses.Add(BaseRange.GetRangeAddress(Worksheet, address));
                }

                return string.Join(",", addresses.ToArray());
            }
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private static void GetAddressElementsFromAddressString(Workbook workbook, string addressString, out Worksheet worksheet, out List<string> localAddresses)
        {
            // DefinedNames can contain multiple (fully qualified) addresses. These are comma separated.
            // Worksheet names containing commas (or spaces) are quoted. Apostrophes in a worksheet name are escaped with a double apostrophe.
            // need to iterate each address and extract just the address part. Names cannot span multiple worksheets, and we are only interested in the address.            

            List<string> addresses = GetAddressesFromAddressString(addressString);

            worksheet = null;
            string worksheetName;
            localAddresses = GetLocalAddresses(addresses, out worksheetName).ToList();
            if (workbook.Worksheets.Contains(worksheetName))
                worksheet = workbook.Worksheets[worksheetName];
        }

        private static List<string> GetAddressesFromAddressString(string addressString)
        {
            // This is a more complicated version of addressString.Split(',').
            // Sheet names can contain commas. These sheet names are then quoted. This method splits allAddresses for non-quoted commas.

            List<string> addresses = new List<string>();

            int currentIndex = 0;
            while (currentIndex >= 0)
            {
                int commaIndex = addressString.IndexOf(",", currentIndex);

                if (commaIndex < 0)
                {
                    // exit loop if no commas found - we only care about finding multiple addresses here.
                    break;
                }
                else
                {
                    int quoteIndex = addressString.IndexOf("'", currentIndex);

                    if (quoteIndex >= 0)
                    {
                        // Get next single apostrophe. Apostrophes are valid sheet name characters when not at the start or end. These are escaped with a double apostrophe.
                        int endQuoteIndex = addressString.IndexOf("'", quoteIndex + 1);

                        while (addressString.Substring(endQuoteIndex + 1, 1) == "'")
                        {
                            endQuoteIndex = addressString.IndexOf("'", endQuoteIndex + 2);
                        }

                        commaIndex = addressString.IndexOf(",", endQuoteIndex);
                    }

                    if (commaIndex < 0)
                        break;

                    // get address and unescape double apostrophes
                    string address = addressString.Substring(currentIndex, commaIndex - currentIndex).Replace("''", "'");
                    addresses.Add(address);

                    currentIndex = commaIndex + 1;
                }
            }

            // add last address
            addresses.Add(addressString.Substring(currentIndex).Replace("''", "'"));

            return addresses;
        }

        private static string[] GetLocalAddresses(List<string> addresses, out string worksheetName)
        {
            // This method takes a list of fully-qualified addresses and returns the addresses minus sheet names.
            // The name of the first worksheet is also returned as worksheetName. Names cannot span multiple worksheets, so this is safe.

            worksheetName = "";

            List<string> localAddresses = new List<string>();

            foreach (string address in addresses)
            {
                int separatorPosition = address.LastIndexOf('!');

                if (worksheetName == "")
                {
                    worksheetName = address.Substring(0, separatorPosition);
                    if (worksheetName.StartsWith("'") && worksheetName.EndsWith("'"))
                        worksheetName = worksheetName.Substring(1, worksheetName.Length - 2);                    
                }

                localAddresses.Add(address.Substring(separatorPosition + 1));                
            }

            return localAddresses.ToArray();
        }

        
        /***********************************
         * DAL METHODS
         ************************************/

        // Read
        internal static DefinedName ReadDefinedNameFromReader(CustomOpenXmlReader reader, Workbook workbook, int nameId)
        {
            DefinedName definedName = new DefinedName(workbook);
            definedName.NameId = nameId;

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "name":
                        definedName.Name = attribute.Value;
                        break;
                    case "localSheetId":
                        definedName.Scope = DefinedNameScope.Worksheet;
                        break;
                }
            }

            definedName.SetAddresses(reader.GetText());

            return definedName;
        }

        // Write

        internal static void AddDefinedNameToDefinedNamesElement(OpenXmlSpreadsheet.DefinedNames definedNamesElement, DefinedName definedName)
        {
            OpenXmlSpreadsheet.DefinedName definedNameElement = new OpenXmlSpreadsheet.DefinedName();
            definedNameElement.Name = definedName.Name;
            
            string addressString = definedName.GetAddressString(false);
            definedNameElement.Text = addressString;
            
            if (definedName.Scope == DefinedNameScope.Worksheet) definedNameElement.LocalSheetId = new OpenXml.UInt32Value((UInt32)definedName.Worksheet.Index - 1);

            definedNamesElement.Append(definedNameElement);
        }

    }
}
