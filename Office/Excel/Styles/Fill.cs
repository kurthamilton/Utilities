using System;
using System.Collections.Generic;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{    

    public class Fill : BaseExcel, IEquatable<Fill>
    {
        internal const int DefaultFillId = 0;
        
        internal Styles Styles { get; private set; }
        private BaseRange BaseRange { get; set; }

        private int _fillId = DefaultFillId;
        internal int FillId { get { return _fillId; } private set { _fillId = value; } }

        private PatternType _patternType = PatternType.Solid;
        public PatternType PatternType { get { return _patternType; } set { UpdateFillProperty(FillProperty.PatternType, value); } }

        private Color _foregroundColor;
        public Color ForegroundColor
        {
            get { return GetForegroundColor(); }
            set { UpdateFillProperty(FillProperty.ForegroundColor, value); }
        }

        private Color _backgroundColor;
        public Color BackgroundColor
        {
            get { return GetBackgroundColor(); }
            set { UpdateFillProperty(FillProperty.BackgroundColor, value); }
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        // range level fill - used to update range fill properties
        internal Fill(BaseRange range)
            : this(range.Worksheet.Workbook.Styles.CellFormats[range.StyleIndex].Fill.Clone(range))
        {
            BaseRange = range;

            if (_foregroundColor != null)
                _foregroundColor.Updated += UpdateForegroundColor;
            if (_backgroundColor != null)
                _backgroundColor.Updated += UpdateBackgroundColor;
        }

        // workbook level fills - used to define workbook fills
        internal Fill(Fill fill)
            : this(fill.Styles, fill.FillId, fill.PatternType, fill.ForegroundColor, fill.BackgroundColor)
        {
        }

        internal Fill(Styles styles)
            : this(styles, DefaultFillId)
        {
        }

        internal Fill(Styles styles, int fillId)
        {
            Styles = styles;
            FillId = fillId;
        }

        internal Fill(Styles styles, int fillId, PatternType patternType, Color foregroundColor, Color backgroundColor)
        {
            Styles = styles;
            FillId = fillId;
            SetFillProperty(FillProperty.PatternType, patternType);
            SetFillProperty(FillProperty.ForegroundColor, foregroundColor);
            SetFillProperty(FillProperty.BackgroundColor, backgroundColor);
        }

        /***********************************
         * PUBLIC METHODS
         ************************************/

        public Fill Clone(BaseRange baseRange)
        {            
            return new Fill(Styles, -1, PatternType, ForegroundColor.Clone(), BackgroundColor.Clone());
        }

        // implement IEquatable
        public bool Equals(Fill other)
        {
            return (other.PatternType == PatternType && 
                other.ForegroundColor.Equals(ForegroundColor) && 
                other.BackgroundColor.Equals(BackgroundColor));
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/


        /***********************************
         * PRIVATE METHODS
         ************************************/

        private Color GetForegroundColor()
        {
            if (_foregroundColor == null)
            {
                if (BaseRange != null)
                {
                    _foregroundColor = new Color(BaseRange);
                    _foregroundColor.Updated += UpdateForegroundColor;
                }
                else
                    _foregroundColor = new Color();
            }
            return _foregroundColor;
        }
        private Color GetBackgroundColor()
        {            
            if (_backgroundColor == null)
            {
                if (BaseRange != null)
                {
                    _backgroundColor = new Color(BaseRange);
                    _backgroundColor.Updated += UpdateBackgroundColor;
                }
                else
                    _backgroundColor = new Color();
            }
            return _backgroundColor;
        }

        private void UpdateForegroundColor(object sender, EventArgs e)
        {
            UpdateFillProperty(FillProperty.ForegroundColor, sender);            
        }
        private void UpdateBackgroundColor(object sender, EventArgs e)
        {
            UpdateFillProperty(FillProperty.BackgroundColor, sender);
        }

        private void SetFillProperty(FillProperty fillProperty, object value)
        {
            if (value != null)
            {
                switch (fillProperty)
                {
                    case FillProperty.PatternType:
                        if (Enum.IsDefined(typeof(PatternType), value))
                            _patternType = (PatternType)value;
                        else
                            return;
                        break;        
                    case FillProperty.BackgroundColor:
                        _backgroundColor = (Color)value;
                        break;
                    case FillProperty.ForegroundColor:
                        _foregroundColor = (Color)value;
                        break;
                    default:
                        throw new Exception(string.Format("FillProperty {0} not implemented in Fill.SetFillProperty", fillProperty));
                }
            }
        }

        private void UpdateFillProperty(FillProperty fillProperty, object value)
        {
            SetFillProperty(fillProperty, value);

            if (fillProperty == FillProperty.ForegroundColor && PatternType == Excel.PatternType.None)
                SetFillProperty(FillProperty.PatternType, Excel.PatternType.Solid);

            if (BaseRange != null)
            {
                Styles styles = BaseRange.Worksheet.Workbook.Styles;

                // update base range with new/existing cell format id
                Fill newFill = styles.Fills.Insert(this);
                CellFormat cellFormat = styles.CellFormats[BaseRange.StyleIndex];
                CellFormat newCellFormat = styles.CellFormats.Insert(
                    new CellFormat(styles, cellFormat.Alignment, cellFormat.Borders, newFill, cellFormat.Font, cellFormat.NumberFormat));
                BaseRange.StyleIndex = newCellFormat.CellFormatId;
            }
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static Fill ReadFillFromReader(CustomOpenXmlReader reader, Styles styles, int fillId)
        {
            Fill fill = new Fill(styles);
            fill.FillId = fillId;

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.Fill>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.PatternFill>())
                    ReadPatternFillFromReader(reader, fill);
            }

            return fill;
        }

        private static void ReadPatternFillFromReader(CustomOpenXmlReader reader, Fill fill)
        {
            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "patternType":
                        fill.PatternType = Helpers.GetEnumValueFromDescription<PatternType>(attribute.Value);
                        break;                        
                }
            }

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.PatternFill>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.ForegroundColor>())
                    fill.ForegroundColor = Color.ReadColorFromReader(reader);
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.BackgroundColor>())
                    fill.BackgroundColor = Color.ReadColorFromReader(reader);
            }
        }

        // Write

        internal static void WriteFillToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Fill fill)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.Fill());
            WritePatternFillToWriter(writer, fill);
            writer.WriteEndElement();   // Fill
        }

        private static void WritePatternFillToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, Fill fill)
        {
            if (fill.PatternType != PatternType.None)
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.PatternFill());

                writer.WriteAttribute("patternType", Helpers.ToCamelCase(fill.PatternType.ToString()));

                if (fill._foregroundColor != null) Color.WriteColorToWriter(writer, fill.ForegroundColor, new OpenXmlSpreadsheet.ForegroundColor());
                if (fill._backgroundColor != null) Color.WriteColorToWriter(writer, fill.BackgroundColor, new OpenXmlSpreadsheet.BackgroundColor());

                writer.WriteEndElement();   // PatternFill
            }
        }
    }
}
