using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    public class NumberFormatCollection : BaseExcel, IEnumerable<NumberFormat>
    {
        private Styles Styles { get; set; }

        private SortedDictionary<int, NumberFormat> _builtInFormats = new SortedDictionary<int, NumberFormat>();
        private SortedDictionary<int, NumberFormat> BuiltInFormats { get { return _builtInFormats; } }

        private SortedDictionary<int, NumberFormat> _numberFormats = new SortedDictionary<int, NumberFormat>();
        private SortedDictionary<int, NumberFormat> NumberFormats { get { return _numberFormats; } }        
        
        public int Count { get { return NumberFormats.Count; } }


        /***********************************
         * CONSTRUCTORS
         ************************************/


        internal NumberFormatCollection(Styles styles)
        {
            Styles = styles;
            AddDefaultNumberFormatToCollection();
            GetBuiltInNumberFormats();            
        }        


        /***********************************
         * PUBLIC METHODS
         ************************************/


        public NumberFormat this[int index]
        {
            get
            {                
                if (BuiltInFormats.ContainsKey(index))
                    return BuiltInFormats[index];
                if (NumberFormats.ContainsKey(index))
                    return NumberFormats[index];

                return null;
            }
        }

        public NumberFormat this[string formatCode]
        {
            get
            {
                NumberFormat numberFormat = GetNumberFormatByFormatCode(BuiltInFormats.Values, formatCode);
                if (numberFormat != null)
                    numberFormat = GetNumberFormatByFormatCode(NumberFormats.Values, formatCode);                
                return numberFormat;
            }
        }

        public NumberFormat this[NumberFormat numberFormat]
        {
            get
            {
                if (numberFormat != null)
                {
                    NumberFormat match = BuiltInFormats.Values.FirstOrDefault(nf => nf.Equals(numberFormat));
                    if (match == null)
                        match = NumberFormats.Values.FirstOrDefault(nf => nf.Equals(numberFormat));
                    return match;
                }

                return null;
            }
        }

        public NumberFormat Insert(NumberFormat numberFormat)
        {
            NumberFormat existingNumberFormat = this[numberFormat];

            if (existingNumberFormat == null)
            {
                int newNumberFormatId = GenerateNewNumberFormatId();
                NumberFormat newNumberFormat = new NumberFormat(numberFormat.Styles, newNumberFormatId, numberFormat.FormatCode);
                AddNumberFormatToCollection(newNumberFormat);
                return newNumberFormat;
            }

            return existingNumberFormat;
        }

        // Implement IEnumerable
        public IEnumerator<NumberFormat> GetEnumerator()
        {
            return new GenericEnumerator<NumberFormat>(NumberFormats);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private int GenerateNewNumberFormatId()
        {
            int newNumberFormatId = 164;
            if (NumberFormats.Count > 1)    // account for default number format, so only get max existing value when > 1
                newNumberFormatId = NumberFormats.Keys.Max() + 1;
            return newNumberFormatId;
        }

        private void AddDefaultNumberFormatToCollection()
        {
            AddNumberFormatToCollection(new NumberFormat(Styles));
        }

        private void GetBuiltInNumberFormats()
        {
            foreach (NumberFormat numberFormat in NumberFormatCollection.GetBuiltInNumberFormats(Styles))
            {
                BuiltInFormats.Add(numberFormat.NumberFormatId, numberFormat);
            }
        }

        private NumberFormat GetNumberFormatByFormatCode(IEnumerable<NumberFormat> numberFormats, string formatCode)
        {
            return numberFormats.FirstOrDefault(nf => nf.FormatCode == formatCode);
        }

        private void AddNumberFormatToCollection(NumberFormat numberFormat)
        {
            if (GetNumberFormatByFormatCode(NumberFormats.Values, numberFormat.FormatCode) == null)
            {
                NumberFormats.Add(numberFormat.NumberFormatId, numberFormat);
            }
        }        

        private static IEnumerable<NumberFormat> GetBuiltInNumberFormats(Styles styles)
        {
            // Number formats are unlike other areas of OpenXml. There are a list of reserved number formats (below) that don't exist in any Xml files.            
            // Reference - http://idippedut.dk/post/2009/01/29/The-complexity-of-SpreadsheetML-oh-the-sheer-joy-of-it!.aspx
            // The full list is in Part 4 of the standards documents (link in BaseOffice class), Page 2128
            // Number format ids < 164 are built in. Only the supported built in number formats are below.

            // Some of the Excel built int formats use the system defaults. Assign current culture default patterns to these.
            DateTimeFormatInfo dateTimeFormat = CultureInfo.CurrentCulture.DateTimeFormat;
            NumberFormatInfo numberFormat = CultureInfo.CurrentCulture.NumberFormat;

            string decimalSeparator = numberFormat.NumberDecimalSeparator;
            string thousandsSeparator = numberFormat.NumberGroupSeparator;

            List<NumberFormat> numberFormats = new List<NumberFormat>();

            numberFormats.Add(new NumberFormat(styles, 0, "", true, NumberFormatType.General, ""));
            numberFormats.Add(new NumberFormat(styles, 1, "0", true, NumberFormatType.Numeric, "0"));
            numberFormats.Add(new NumberFormat(styles, 2, "0.00", true, NumberFormatType.Numeric, string.Format("0{0}00", decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 3, "#,##0", true, NumberFormatType.Numeric, string.Format("#{0}##0", thousandsSeparator)));
            numberFormats.Add(new NumberFormat(styles, 4, "#,##0.00", true, NumberFormatType.Numeric, string.Format("#{0}##0{1}00", thousandsSeparator, decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 9, "0%", true, NumberFormatType.Numeric, "0%"));
            numberFormats.Add(new NumberFormat(styles, 10, "0.00%", true, NumberFormatType.Numeric, string.Format("0{0}00%", decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 11, "0.00E+00", true, NumberFormatType.Numeric, string.Format("0{0}00E00", decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 12, "# ?/?", true));
            numberFormats.Add(new NumberFormat(styles, 13, "# ??/??", true));
            numberFormats.Add(new NumberFormat(styles, 14, "mm-dd-yy", true, NumberFormatType.DateTime, dateTimeFormat.ShortDatePattern));
            numberFormats.Add(new NumberFormat(styles, 15, "d-mmm-yy", true, NumberFormatType.DateTime, "d-MMM-yy"));
            numberFormats.Add(new NumberFormat(styles, 16, "d-mmm", true, NumberFormatType.DateTime, "d-MMM"));
            numberFormats.Add(new NumberFormat(styles, 17, "mmm-yy", true, NumberFormatType.DateTime, "MMM-yy"));
            numberFormats.Add(new NumberFormat(styles, 18, "h:mm AM/PM", true, NumberFormatType.DateTime, dateTimeFormat.ShortTimePattern));  // don't support AM/PM "h:mm tt"
            numberFormats.Add(new NumberFormat(styles, 19, "h:mm:ss AM/PM", true, NumberFormatType.DateTime, dateTimeFormat.LongTimePattern));
            numberFormats.Add(new NumberFormat(styles, 20, "h:mm", true, NumberFormatType.DateTime, dateTimeFormat.ShortTimePattern));
            numberFormats.Add(new NumberFormat(styles, 21, "h:mm:ss", true, NumberFormatType.DateTime, dateTimeFormat.LongTimePattern));
            numberFormats.Add(new NumberFormat(styles, 22, "m/d/yy h:mm", true, NumberFormatType.DateTime, string.Concat(dateTimeFormat.ShortDatePattern, " ", dateTimeFormat.LongTimePattern)));
            numberFormats.Add(new NumberFormat(styles, 37, "#,##0 ;(#,##0)", true, NumberFormatType.Numeric, string.Format("#{0}##0;(#{0}##0)", thousandsSeparator)));
            numberFormats.Add(new NumberFormat(styles, 38, "#,##0 ;[Red](#,##0)", true, NumberFormatType.Numeric, string.Format("#{0}##0;(#{0}##0)", thousandsSeparator)));
            numberFormats.Add(new NumberFormat(styles, 39, "#,##0.00 ;(#,##0.00)", true, NumberFormatType.Numeric, string.Format("#{0}##0{1}00;(#{0}##0{1}00)", thousandsSeparator, decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 40, "#,##0.00 ;[Red](#,##0.00)", true, NumberFormatType.Numeric, string.Format("#{0}##0{1}00;(#{0}##0{1}00)", thousandsSeparator, decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 45, "mm:ss", true, NumberFormatType.DateTime, "mm:ss"));
            numberFormats.Add(new NumberFormat(styles, 46, "[h]:mm:ss", true, NumberFormatType.DateTime));
            numberFormats.Add(new NumberFormat(styles, 47, "mmss.0", true, NumberFormatType.DateTime, "mmss.f"));
            numberFormats.Add(new NumberFormat(styles, 48, "##0.0E+0", true, NumberFormatType.Numeric, string.Format("##0{0}0E0", decimalSeparator)));
            numberFormats.Add(new NumberFormat(styles, 49, "@", true, NumberFormatType.Text, ""));

            return numberFormats;
        }

        /***********************************
         * DAL METHODS
         ************************************/

        // Read

        internal static NumberFormatCollection ReadNumberFormatsFromReader(CustomOpenXmlReader reader, Styles styles)
        {
            NumberFormatCollection numberFormats = new NumberFormatCollection(styles);

            while (reader.ReadToEndElement<OpenXmlSpreadsheet.NumberingFormats>())
            {
                if (reader.IsStartElementOfType<OpenXmlSpreadsheet.NumberingFormat>())
                {
                    NumberFormat numberFormat = NumberFormat.ReadNumberFormatFromReader(reader, styles, numberFormats.NumberFormats.Keys.Max() + 1);
                    numberFormats.AddNumberFormatToCollection(numberFormat);
                }
            }

            return numberFormats;
        }

        // Write

        internal static void WriteNumberFormatsToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, IEnumerable<NumberFormat> numberFormats)
        {
            writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.NumberingFormats());

            foreach (NumberFormat numberFormat in numberFormats)
            {
                NumberFormat.WriteNumberFormatToWriter(writer, numberFormat);
            }

            writer.WriteEndElement();   // NumberingFormats
        }

    }
}
