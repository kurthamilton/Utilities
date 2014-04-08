using System;
using System.Collections.Generic;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal enum FontProperty : int
    {
        Size,
        Color,
        Name,
        Bold,
        Italic,
        Underline
    }

    public class Font : BaseExcel, IEquatable<Font>
    {
        internal const int DefaultFontId = 0;
        internal List<FontProperty> UsedFontProperties = new List<FontProperty>();

        internal Styles Styles { get; private set; }
        private BaseRange BaseRange { get; set; }

        private int _fontId = DefaultFontId;
        internal int FontId { get { return _fontId; } private set { _fontId = value; } }

        // need to update Equals method if adding properties that are classes
        private double _size;
        public double Size { get { return _size; } set { UpdateFontProperty(FontProperty.Size, value); } }
        private Color _color;
        public Color Color 
        { 
            get { return GetColor(); } 
            set { UpdateFontProperty(FontProperty.Color, value); } 
        }
        private string _name;
        public string Name { get { return _name; } set { UpdateFontProperty(FontProperty.Name, value); } }
        private bool _bold;
        public bool Bold { get { return _bold; } set { UpdateFontProperty(FontProperty.Bold, value); } }
        private bool _italic;
        public bool Italic { get { return _italic; } set { UpdateFontProperty(FontProperty.Italic, value); } }
        private bool _underline;
        public bool Underline { get { return _underline; } set { UpdateFontProperty(FontProperty.Underline, value); } }

        /***********************************
         * CONSTRUCTORS
         ************************************/

        public Font()
        {
        }

        // range level font - used to update range font properties
        internal Font(BaseRange range)
            : this(range.Worksheet.Workbook.Styles.CellFormats[range.StyleIndex].Font)
        {
            BaseRange = range;

            if (_color != null)
                _color.Updated += UpdateColor;
        }

        // workbook level fonts - used to define workbook fonts
        internal Font(Font font)
            : this(font.Styles, font.FontId, font.GetFontProperties())
        {
        }

        internal Font(Styles styles, Dictionary<FontProperty, object> fontProperties)
            : this(styles)
        {
        }

        internal Font(Styles styles)
            : this(styles, DefaultFontId)
        {
        }

        internal Font(Styles styles, int fontId)
            : this(styles, fontId, null)
        {
        }

        internal Font(Styles styles, int fontId, Dictionary<FontProperty, object> fontProperties)
        {
            Styles = styles;
            FontId = fontId;

            if (fontProperties != null)
            {
                foreach (KeyValuePair<FontProperty, object> fontProperty in fontProperties)
                {
                    SetFontProperty(fontProperty.Key, fontProperty.Value);
                }
            }
        }

        /***********************************
         * INTERNAL PROPERTIES
         ************************************/

        // Character width and height calculations aren't based on the documentation, instead I have ascertained a common formula using different input values
        // by trial and error to approximately match real values. 
        // The actual calculations for column width and row height are a function of both font size and cell text length, 
        // so it should in fact be more efficient to store the static char widths and heights with the font, although obviously with loss of accuracy.
        private double _characterWidth = -1;
        internal double CharacterWidth
        {
            get
            {
                if (_characterWidth == -1)
                {
                    // I'm not sure this is exactly right. The default font seems to be stored at index 1, but cells are using index 0 by default. Need to investigate.
                    // This is a quick fix to ensure the CharacterWidth is positive.
                    double size = Size;
                    if (size <= 0)
                        size = Styles.Fonts[1].Size;
                    _characterWidth = GetCharacterWidth(size);
                }
                return _characterWidth;
            }
        }
        private double _characterHeight = -1;
        internal double CharacterHeight
        {
            get
            {
                if (_characterHeight == -1)
                {
                    // I'm not sure this is exactly right. The default font seems to be stored at index 1, but cells are using index 0 by default. Need to investigate.
                    // This is a quick fix to ensure the CharacterHeight is positive.
                    double size = Size;
                    if (size <= 0)
                        size = Styles.Fonts[1].Size;
                    _characterHeight = GetCharacterHeight(size);
                }
                return _characterHeight;
            }
        }


        /***********************************
         * PUBLIC METHODS
         ************************************/

        public Font Clone()
        {
            return new Font(this);
        }

        // implement IEquatable
        public bool Equals(Font other)
        {
            Dictionary<FontProperty, object> fontProperties = GetFontProperties();
            Dictionary<FontProperty, object> otherFontProperties = other.GetFontProperties();

            if (fontProperties.Count == otherFontProperties.Count)
            {
                foreach (KeyValuePair<FontProperty, object> property in fontProperties)
                {
                    if (otherFontProperties.ContainsKey(property.Key))
                    {                        
                        // calling .Equals on an object doesn't call the type specific implementation, so need to explicitly call it.
                        if (property.Value.GetType() == typeof(Color))
                        {
                            if (!((Color)property.Value).Equals((Color)otherFontProperties[property.Key]))
                                return false;
                        }
                        else if (!property.Value.Equals(otherFontProperties[property.Key]))
                            return false;
                    }
                    else
                        return false;
                }
                return true;
            }

            return false;
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal object GetFontProperty(FontProperty fontProperty)
        {
            if (UsedFontProperties.Contains(fontProperty))
            {
                switch (fontProperty)
                {
                    case FontProperty.Size:
                        return _size;
                    case FontProperty.Color:
                        return _color;
                    case FontProperty.Name:
                        return _name;
                    case FontProperty.Bold:
                        return _bold;
                    case FontProperty.Italic:
                        return _italic;
                    case FontProperty.Underline:
                        return _underline;
                    default:
                        throw new Exception(string.Format("FontProperty {0} not implemented in Font.GetFontProperty", fontProperty));
                }
            }
            else
                return null;
        }

        internal Dictionary<FontProperty, object> GetFontProperties()
        {
            Dictionary<FontProperty, object> fontProperties = new Dictionary<FontProperty, object>();

            foreach (FontProperty fontProperty in UsedFontProperties)
            {
                fontProperties.Add(fontProperty, GetFontProperty(fontProperty));
            }

            return fontProperties;
        }

        internal static double GetCharacterWidth(double fontSize)
        {
            return CalculateCharacterDimension(fontSize, 5, 20, 1.35);
        }

        internal static double GetCharacterHeight(double fontSize)
        {
            return CalculateCharacterDimension(fontSize, 10, 2.5, 1.35);
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private Color GetColor()
        {
            if (_color == null)
            {
                if (BaseRange != null)
                {
                    _color = new Color(BaseRange);
                    _color.Updated += UpdateColor;
                }
                else
                    _color = new Color();
            }
            return _color;
        }

        private void UpdateColor(object sender, EventArgs e)
        {
            UpdateFontProperty(FontProperty.Color, sender);
        }

        private void SetFontProperty(FontProperty fontProperty, object value)
        {
            if (value != null)
            {
                switch (fontProperty)
                {
                    case FontProperty.Size:
                        if ((double)value > 0)
                            _size = (double)value;
                        else
                            return;
                        break;
                    case FontProperty.Color:
                        Color color = (Color)value;
                        if (color.IsUsed())                        
                            _color = color;
                        else
                            return;
                        break;
                    case FontProperty.Name:
                        if (!string.IsNullOrEmpty(value.ToString()))
                            _name = value.ToString();
                        else
                            return;
                        break;
                    case FontProperty.Bold:
                        _bold = (bool)value;
                        break;
                    case FontProperty.Italic:
                        _italic = (bool)value;
                        break;
                    case FontProperty.Underline:
                        _underline = (bool)value;
                        break;
                    default:
                        throw new Exception(string.Format("FontProperty {0} not implemented in Font.UpdateFontProperty", fontProperty));
                }
                if (!UsedFontProperties.Contains(fontProperty))
                    UsedFontProperties.Add(fontProperty);
            }
            else
                UsedFontProperties.Remove(fontProperty);
        }

        private void UpdateFontProperty(FontProperty fontProperty, object value)
        {
            SetFontProperty(fontProperty, value);

            if (BaseRange != null)
            {
                // update base range with new/existing cell format id
                Font newFont = BaseRange.Worksheet.Workbook.Styles.Fonts.Insert(this);
                CellFormat cellFormat = BaseRange.Worksheet.Workbook.Styles.CellFormats[BaseRange.StyleIndex];
                CellFormat newCellFormat = BaseRange.Worksheet.Workbook.Styles.CellFormats.Insert(
                    new CellFormat(cellFormat.Styles, cellFormat.Alignment, cellFormat.Borders, cellFormat.Fill, newFont, cellFormat.NumberFormat));
                BaseRange.StyleIndex = newCellFormat.CellFormatId;
            }
        }


        private static double CalculateCharacterDimension(double fontSize, double topValue, double bottomValue, double power)
        {
            if (fontSize > 0)
                return (topValue + Math.Pow(fontSize, power)) / (bottomValue);
            else
                return 0;
        }

        /***********************************
        * DAL METHODS
        ************************************/

        // Read

        internal static void UpdateFontFromReader<TFont>(CustomOpenXmlReader reader, Font font) where TFont : OpenXml.OpenXmlElement
        {
            // Font properties can either be stored in a Font element, or a Run Properties element (in the shared string table). 
            // They use mainly similar child elements, apart from where specified below.

            while (reader.ReadToEndElement<TFont>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.FontSize>())
                    font.Size = reader.Attributes["val"].GetDoubleValue();
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Color>())
                    font.Color = Color.ReadColorFromReader(reader);
                if ((typeof(TFont) == typeof(OpenXmlSpreadsheet.Font) && reader.IsStartElementOfType<OpenXmlSpreadsheet.FontName>()) ||
                    (typeof(TFont) == typeof(OpenXmlSpreadsheet.RunProperties) && reader.IsStartElementOfType<OpenXmlSpreadsheet.RunFont>()))
                    font.Name = reader.Attributes["val"].Value;
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Bold>())
                    font.Bold = true;
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Italic>())
                    font.Italic = true;
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Underline>())
                    font.Underline = true;
            }
        }

        internal static Font ReadFontFromFontElement(CustomOpenXmlReader reader, Styles styles, int fontId)
        {
            Font font = new Font(styles);
            font.FontId = fontId;

            UpdateFontFromReader<OpenXmlSpreadsheet.Font>(reader, font);

            return font;
        }

        internal static Font ReadFontFromRunPropertiesElement(CustomOpenXmlReader reader)
        {
            Font font = new Font();

            UpdateFontFromReader<OpenXmlSpreadsheet.RunProperties>(reader, font);

            return font;
        }

        // Write

        internal static void WriteFontElement(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Font font)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Font());

            WriteFontToWriter<OpenXmlPackaging.WorkbookStylesPart>(writer, font);

            writer.WriteEndElement();   // Font
        }

        internal static void WriteRunPropertiesElement(CustomOpenXmlWriter<OpenXmlPackaging.SharedStringTablePart> writer, Font font)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.RunProperties());

            WriteFontToWriter<OpenXmlPackaging.SharedStringTablePart>(writer, font);

            writer.WriteEndElement();   // RunProperties
        }

        private static void WriteFontToWriter<T>(CustomOpenXmlWriter<T> writer, Font font) where T : OpenXmlPackaging.OpenXmlPart, OpenXmlPackaging.IFixedContentTypePart
        {
            if (font.Size > 0)
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.FontSize());
                writer.WriteAttribute("val", font.Size);
                writer.WriteEndElement();   // FontSize
            }

            Color.WriteColorToWriter(writer, font.Color);

            if (!string.IsNullOrEmpty(font.Name))
            {
                if (typeof(T) == typeof(OpenXmlPackaging.WorkbookStylesPart))
                    writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.FontName());
                else if (typeof(T) == typeof(OpenXmlPackaging.SharedStringTablePart))
                    writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.RunFont());

                writer.WriteAttribute("val", font.Name);
                writer.WriteEndElement();   // FontName / RunFont
            }

            if (font.Bold) writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Bold(), true);
            if (font.Italic) writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Italic(), true);
            if (font.Underline) writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Underline(), true);

            
        }
        
    }
}
