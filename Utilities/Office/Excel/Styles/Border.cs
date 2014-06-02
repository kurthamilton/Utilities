using System;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    

    public class Border : BaseExcel, IEquatable<Border>
    {
        internal Borders Borders { get; private set; }
        
        internal BorderType Type { get; private set; }

        private BorderStyle _borderStyle = BorderStyle.None;
        public BorderStyle BorderStyle { get { return _borderStyle; } set { UpdateBorderProperty(BorderProperty.BorderStyle, value); } }

        // colour is not publicly supported at the moment, as it's not very crash-proof. Only used to persist existing colours for the time being.
        internal Color Color { get { if (_color == null) _color = new Color(); return _color; } set { UpdateBorderProperty(BorderProperty.Color, value); } }
        private Color _color;
                
        internal Border(Borders borders, Border border)
            : this(borders, border.Type, border.BorderStyle, border.Color)
        {
        }

        internal Border(Borders borders, BorderType type)
            : this(borders, type, BorderStyle.None, null)
        {
        }

        internal Border(Borders borders, BorderType type, BorderStyle borderStyle, Color color)
        {
            Borders = borders;
            Type = type;
            SetBorderProperty(BorderProperty.BorderStyle, borderStyle);
            SetBorderProperty(BorderProperty.Color, color);
        }

        /***********************************
         * INTERNAL PROPERTIES
         ************************************/


        /***********************************
         * PUBLIC METHODS
         ************************************/

        // implement IEquatable
        public bool Equals(Border other)
        {
            return (other.BorderStyle == BorderStyle && 
                other.Color.Equals(Color));
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal Border Clone()
        {
            return new Border(this.Borders, this);
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void SetBorderProperty(BorderProperty borderProperty, object value)
        {
            if (value != null)
            {
                switch (borderProperty)
                {
                    case BorderProperty.BorderStyle:
                        if (Enum.IsDefined(typeof(BorderStyle), value))
                            _borderStyle = (BorderStyle)value;
                        else
                            return;
                        break; 
                    case BorderProperty.Color:
                        _color = (Color)value;
                        break;
                    default:
                        throw new Exception(string.Format("BorderProperty {0} not implemented in Border.SetBorderProperty", borderProperty));
                }
            }
        }

        private void UpdateBorderProperty(BorderProperty borderProperty, object value)
        {
            SetBorderProperty(borderProperty, value);
            Borders.UpdateBorder(this);
        }


        /***********************************
         * DAL METHODS
         ************************************/

        // Read
        internal static Border ReadBorderFromReader(CustomOpenXmlReader reader, Borders borders)
        {
            BorderType type = Helpers.GetEnumValueFromDescription<BorderType>(reader.LocalName);

            Border border = new Border(borders, type);

            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "style":
                        border.BorderStyle = Helpers.GetEnumValueFromDescription<BorderStyle>(attribute.Value);
                        break;
                }
            }

            Type borderElementType = reader.ElementType;
            while (reader.ReadToEndElement(borderElementType))
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.Color>())
                    border.Color = Color.ReadColorFromReader(reader);
            }

            return border;
        }

        // Write
        internal static void WriteBorderToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Border border, OpenXml.OpenXmlElement borderElement)
        {
            writer.WriteOpenXmlElement(borderElement);

            if (border.BorderStyle != BorderStyle.None)
            {
                writer.WriteAttribute("style", Helpers.ToCamelCase(border.BorderStyle.ToString()));

                if (border._color != null) Color.WriteColorToWriter(writer, border.Color);
            }

            writer.WriteEndElement();   // borderElement
        }
    }
}
