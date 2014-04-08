using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal class FontCollection : BaseExcel, IEnumerable<Font>
    {
        private Styles Styles { get; set; }

        private SortedDictionary<int, Font> _fontDictionary;
        private SortedDictionary<int, Font> FontDictionary { get { if (_fontDictionary == null) _fontDictionary = new SortedDictionary<int,Font>(); return _fontDictionary; } }
        
        public int Count { get { return FontDictionary.Count; } }

        /***********************************
        * CONSTRUCTORS
        ************************************/ 

        public FontCollection(Styles styles)
        {
            Styles = styles;
        }               

        /***********************************
        * PUBLIC PROPERTIES
        ************************************/ 

        public Font DefaultFont
        {
            get
            {
                if (FontDictionary.Count > 0)
                    return FontDictionary[0];
                else
                    return null;
            }
        }

        /***********************************
        * PUBLIC METHODS
        ************************************/ 

        public Font this[int index]
        {
            get
            {
                if (Contains(index))
                    return FontDictionary[index];
                else
                    return null;
            }
        }

        public Font this[Font font]
        {
            get
            {
                if (font != null)
                    return (FontDictionary.Values.FirstOrDefault(f => f.Equals(font)));
                
                return null;
            }
        }

        public void Clear()
        {
            // first font is default font. Don't clear.
            while (FontDictionary.Count > 1)
            {
                FontDictionary.Remove(FontDictionary.Max(f => f.Key));
            }
        }

        public bool Contains(int fontId)
        {
            return FontDictionary.ContainsKey(fontId);
        }

        public Font Insert(Font font)
        {
            Font existingFont = this[font];

            if (existingFont == null)
            {
                int newFontId = GenerateNewFontId();
                Font newFont = new Font(font.Styles, newFontId, font.GetFontProperties());
                AddFontToCollection(newFont);
                return newFont;
            }

            return existingFont;
        }

        public void Delete(int index)
        {
            if (Contains(index))
                FontDictionary.Remove(index);
        }

        // Implement IEnumerable
        public IEnumerator<Font> GetEnumerator()
        {
            return new GenericEnumerator<Font>(FontDictionary);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        

        /***********************************
        * PRIVATE METHODS
        ************************************/

        private int GenerateNewFontId()
        {
            return OfficeUtilities.GetFirstUnusedKeyFromCollection<Font>(FontDictionary);
        }

        private void AddFontToCollection(Font font)
        {
            if (!Contains(font.FontId))
                FontDictionary.Add(font.FontId, font);
        }

        /***********************************
        * DAL METHODS
        ************************************/

        // Read

        internal static FontCollection ReadFontsFromReader(CustomOpenXmlReader reader, Styles styles)
        {
            FontCollection fonts = new FontCollection(styles);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Fonts>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Font>())
                {
                    Font font = Font.ReadFontFromFontElement(reader, styles, fonts.Count + Font.DefaultFontId);
                    fonts.AddFontToCollection(font);
                }
            }

            return fonts;
        }

        // Write

        internal static void WriteFontsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, IEnumerable<Font> fonts)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Fonts());

            foreach (Font font in fonts)
            {
                Font.WriteFontElement(writer, font);
            }

            writer.WriteEndElement();   // Fonts
        }

    }
}
