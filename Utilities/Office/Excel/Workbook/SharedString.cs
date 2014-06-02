using System;
using System.Collections.Generic;
using System.Text;

using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal class SharedString : BaseExcel, IEquatable<SharedString>
    {
        public Workbook Workbook { get; private set; }

        private int _index = -1;
        public int Index { get { return _index; } set { if (_index == -1 && value >= 0) _index = value; } }

        private string Value { get; set; }

        // only initialised and used if shared string has multiple fonts.
        private List<Font> Fonts { get; set; }  
        private List<string> Values { get; set; } // Value is updated with the concatenated version of this list when initialised

        public bool IsFontString { get { return Fonts != null; } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        internal SharedString(Workbook workbook, string value)
        {
            Workbook = workbook;
            Value = value;
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        public void AddFontString(Font font, string value)
        {
            InitialiseForFontStrings();

            Fonts.Add(font);
            Values.Add(value);

            UpdateValue();
        }

        // This is a very simple handling of potentially complicated shared strings. If just dealing with simple strings, there is no problem, but dealing with
        // multiple fonts could get complicated. For example, if a string of abc needs to be added, but a<i>b</i>c already exists, then the cell will use the existing styled abc.
        // Assume that this won't happen, or at least don't worry about the consequences for now.
        public override string ToString()
        {
            return Value;
        }

        // Implement IEquatable
        public bool Equals(SharedString other)
        {
            if (Values == null && other.Values == null)
                return Value == other.Value;

            if (Values != null && other.Values != null && Fonts != null && other.Fonts != null)
            {
                if (Values.Count == other.Values.Count)
                {
                    int i = 0;
                    foreach (string value in Values)
                    {
                        if (value != other.Values[i])
                            return false;

                        i++;
                    }
                }

                if (Fonts.Count == other.Fonts.Count)
                {
                    int i = 0;
                    foreach (Font font in Fonts)
                    {
                        if ((font == null && other.Fonts[i] != null) || (font != null && other.Fonts[i] == null))
                            return false;
                        else if (font != null && other.Fonts[i] != null)
                        {
                            if (!font.Equals(other.Fonts[i]))
                                return false;
                        }
                        i++;
                    }
                }

                return (Values.Count == other.Values.Count && Fonts.Count == other.Fonts.Count);
            }
            
            return false;            
        }


        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void InitialiseForFontStrings()
        {
            if (Fonts == null)
            {
                Fonts = new List<Font>();
                Values = new List<string>();

                if (Value != "")
                {
                    Fonts.Add(null);
                    Values.Add(Value);
                }
            }
        }

        private void UpdateValue()
        {
            if (Values != null)
                Value = string.Join("", Values);
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read
        internal static SharedString ReadSharedStringFromReader(CustomOpenXmlReader reader, Workbook workbook)
        {
            if (reader.ReadFirstChild())
            {                
                if (reader.ElementType == typeof(OpenXmlSpreadsheet.Text))
                {
                    // simple string

                    string value = reader.GetText();
                    SharedString sharedString = new SharedString(workbook, value);
                    return sharedString;
                }
                else if (reader.ElementType == typeof(OpenXmlSpreadsheet.Run))
                {
                    // multiple font string

                    SharedString sharedString = new SharedString(workbook, "");
                    
                    // We have already moved to the Run element by having to determine the element type. Add the first font value outside of the loop
                    AddFontStringFromReader(reader, sharedString);

                    while (reader.ReadToEndElement<OpenXmlSpreadsheet.SharedStringItem>())
                    {
                        if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Run>())
                            AddFontStringFromReader(reader, sharedString);
                    }

                    return sharedString;
                }

            }
            return new SharedString(workbook, "");
        }

        private static void AddFontStringFromReader(CustomOpenXmlReader reader, SharedString sharedString)
        {
            Font font = null;
            string value = "";
            
            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Run>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Text>())
                    value = reader.GetText();
                else if (reader.IsStartElementOfType<OpenXmlSpreadsheet.RunProperties>())
                    font = Font.ReadFontFromRunPropertiesElement(reader);
            }

            sharedString.AddFontString(font, value);
        }

        // Write
        internal static void WriteSharedStringToWriter(CustomOpenXmlWriter<OpenXmlPackaging.SharedStringTablePart> writer, SharedString sharedString)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.SharedStringItem());

            if (sharedString.Values == null)
            {
                // simple string
                WriteTextElementToWriter(writer, sharedString.Value);
            }
            else
            {
                // multiple font string

                for (int i = 0; i < sharedString.Values.Count; i++)
                {
                    writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Run());

                    if (sharedString.Fonts[i] != null)
                    {
                        Font.WriteRunPropertiesElement(writer, sharedString.Fonts[i]);
                    }
                    if (sharedString.Values[i] != null)
                    {
                        WriteTextElementToWriter(writer, sharedString.Values[i]);
                    }

                    writer.WriteEndElement();   // Run
                }
            }

            writer.WriteEndElement(); // SharedStringItem
        }

        private static void WriteTextElementToWriter(CustomOpenXmlWriter<OpenXmlPackaging.SharedStringTablePart> writer, string value)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Text());
            if (value.StartsWith(" ") || value.EndsWith(" "))
                writer.WriteAttribute("space", "preserve", "xml");
            writer.WriteString(value);
            writer.WriteEndElement();    // Text
        }        
    }
}
