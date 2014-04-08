using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

using System.Xml;
using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{
    internal enum NumberFormatProperty : int
    {
        FormatCode
    }

    public class NumberFormat : BaseExcel, IEquatable<NumberFormat>
    {

        // Currently supported number format types:
        //  - General
        //  - Numeric (Decimal, Percentage, Scientific)
        //  - DateTime (simplified)

        internal Styles Styles { get; private set; }
        private BaseRange BaseRange { get; set; }

        internal const int DefaultNumberFormatId = -1;

        internal int NumberFormatId { get; private set; }
        private string _formatCode;
        public string FormatCode { get { return _formatCode; } set { if (!IsBuiltIn) { UpdateNumberFormatProperty(NumberFormatProperty.FormatCode, value); } } }
        internal bool IsBuiltIn = false;
        public NumberFormatType FormatType { get; private set; }
        internal string DotNetFormatCode { get; private set; }        

        string DecimalSeparator;
        string ThousandsSeparator;
        string CurrencySymbol;        

        /***********************************
         * CONSTRUCTORS
         ************************************/

        // range level number format - used to update range number format properties
        internal NumberFormat(BaseRange range)        
        {
            BaseRange = range;
        }

        // workbook level number format - used to define workbook number formats
        internal NumberFormat(Styles styles)
            : this (styles, DefaultNumberFormatId, "", false, NumberFormatType.General, "")
        {
        }

        internal NumberFormat(Styles styles, int numberFormatId, string formatCode)
        {
            Styles = styles;
            NumberFormatId = numberFormatId;
            _formatCode = formatCode;

            NumberFormatInfo numberFormat = CultureInfo.CurrentCulture.NumberFormat;
            DecimalSeparator = numberFormat.NumberDecimalSeparator;
            ThousandsSeparator = numberFormat.NumberGroupSeparator;
            CurrencySymbol = numberFormat.CurrencySymbol;
        }

        internal NumberFormat(Styles styles, int numberFormatId, string formatCode, bool isBuiltIn, 
            NumberFormatType formatType = NumberFormatType.None, string dotNetFormatCode = "")
            : this (styles, numberFormatId, formatCode)
        {
            IsBuiltIn = isBuiltIn;
            FormatType = formatType;
            DotNetFormatCode = dotNetFormatCode;
        }        

        /***********************************
         * PUBLIC METHODS
         ************************************/

        public string GetFormattedValue(object value)
        {
            if (FormatType == NumberFormatType.None)
                SetFormatCode(FormatCode);

            switch (FormatType)
            {
                case NumberFormatType.General:
                    return value.ToString();
                case NumberFormatType.Text:
                    return value.ToString();
                case NumberFormatType.Numeric:
                    return string.Format("{0:" + DotNetFormatCode + "}", value);
                case NumberFormatType.Fraction:
                    throw new NotImplementedException();
                case NumberFormatType.DateTime:
                    {
                        double doubleValue = 0;
                        if (double.TryParse(value.ToString(), out doubleValue))
                        {
                            if (doubleValue >= 0)
                            {
                                DateTime dateTimeValue = ExcelUtilities.GetTimeConverter().GetDateTimeValueFromDoubleValue(doubleValue);
                                if (dateTimeValue > DateTime.MinValue)
                                    return string.Format("{0:" + DotNetFormatCode + "}", dateTimeValue);
                            }
                        }

                        return value.ToString();
                    }
                default:
                    throw new Exception("Invalid Number Format Type");
            }
        }

        // implement IEquatable
        public bool Equals(NumberFormat other)
        {
            return (other.FormatCode == FormatCode);
        }

        

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void SetNumberFormatProperty(NumberFormatProperty numberFormatProperty, object value)
        {
            if (value != null)
            {
                switch (numberFormatProperty)
                {
                    case NumberFormatProperty.FormatCode:
                        SetFormatCode(value.ToString());
                        break;
                    default:
                        throw new Exception(string.Format("NumberFormatProperty {0} not implemented in NumberFormat.UpdateNumberFormatProperty", numberFormatProperty));
                }                
            }
        }

        private void UpdateNumberFormatProperty(NumberFormatProperty numberFormatProperty, object value)
        {
            SetNumberFormatProperty(numberFormatProperty, value);

            if (BaseRange != null)
            {
                // update base range with new/existing cell format id
                NumberFormat newNumberFormat = BaseRange.Worksheet.Workbook.Styles.NumberFormats.Insert(this);
                CellFormat cellFormat = BaseRange.Worksheet.Workbook.Styles.CellFormats[BaseRange.StyleIndex];
                CellFormat newCellFormat = BaseRange.Worksheet.Workbook.Styles.CellFormats.Insert(
                    new CellFormat(cellFormat.Styles, cellFormat.Alignment, cellFormat.Borders, cellFormat.Fill, cellFormat.Font, newNumberFormat));
                BaseRange.StyleIndex = newCellFormat.CellFormatId;
            }
        }

        private void SetFormatCode(string formatCode)
        {
            _formatCode = formatCode;

            // reset properties. These are set accordingly in the TrySetFormats below.
            DotNetFormatCode = "";
            FormatType = NumberFormatType.General;            

            // Excel formats can be split into 4 component parts: Positive;Negative;others..
            // Only take the Positive component to reduce complication
            formatCode = formatCode.Split(';')[0];

            if (TrySetNumericFormat(formatCode))
                return;

            if (TrySetDateTimeFormat(formatCode))
                return;
        }

        private bool TrySetNumericFormat(string formatCode)
        {
            // possible numeric formats:
            // [$-123]  - Currency format code.
            // "£"      - Currency format for system default. £ is replaced by whatever the default is.
            // #,##     - Thousands separator to be used.
            // 0        - Required
            // .00      - Number of decimal places.
            // [$-123]  - Currency format code.
            // E+00     - Scientfic notation 
            // %        - Percentage

            // To do - should add some code to handle leading zeroes ie Excel formats like 000#

            // Thousands and decimal separators are hard coded as comma and point, even if they are set differently in the OS or App.
            // I can't find where custom separators are stored in the XML (maybe they're not - maybe they're in the app settings), 
            // so use default separators when reading data.
            Regex currencyCodeRegex = new Regex(@"(\\\s+)?\[\$.+-[\d\w]+\](\\\s+)?");
            Regex numericRegex = new Regex(string.Format(@"(?<c1>{0})?(?<c2>""{1}"")?(?<t>#,##)?0(?<d>\.0+)?(?<c3>{0})?(?<e>E\+0+)?(?<p>%)?", currencyCodeRegex.ToString(), CurrencySymbol));
            Match numericMatch = numericRegex.Match(formatCode);

            if (numericMatch.Groups["c1"].Success || numericMatch.Groups["c2"].Success || numericMatch.Groups["t"].Success ||
                numericMatch.Groups["d"].Success || numericMatch.Groups["c3"].Success || numericMatch.Groups["e"].Success || numericMatch.Groups["p"].Success)
            {
                StringBuilder dotNetFormatCode = new StringBuilder("");

                FormatType = NumberFormatType.Numeric;                
                
                // the dot net format codes really should be used, such as N2, C2 etc, but that gets a bit more complicated, so just convert each group as it's found

                if (numericMatch.Groups["c1"].Success)
                    dotNetFormatCode.Append(GetCurrencySymbolFromExcelCurrencyCode(numericMatch.Groups["c1"].Value));
                if (numericMatch.Groups["c2"].Success)
                    dotNetFormatCode.Append(CurrencySymbol);
                if (numericMatch.Groups["t"].Success)
                    dotNetFormatCode.Append(numericMatch.Groups["t"].Value.Replace(",", ThousandsSeparator));

                dotNetFormatCode.Append("0");

                if (numericMatch.Groups["d"].Success)
                    dotNetFormatCode.Append(numericMatch.Groups["d"].Value.Replace(".", DecimalSeparator));
                if (numericMatch.Groups["c3"].Success)
                    dotNetFormatCode.Append(GetCurrencySymbolFromExcelCurrencyCode(numericMatch.Groups["c3"].Value));
                if (numericMatch.Groups["e"].Success)
                    dotNetFormatCode.Append(numericMatch.Groups["e"].Value.Replace("+", ""));
                if (numericMatch.Groups["p"].Success)
                    dotNetFormatCode.Append(numericMatch.Groups["p"].Value);
                
                DotNetFormatCode = dotNetFormatCode.ToString();

                return true;
            }

            return false;
        }

        private string GetCurrencySymbolFromExcelCurrencyCode(string currencyCode)
        {
            // Excel currency codes are like
            // <white space>[$<symbol>-<number code>]<white space>
            // This method extracts the all-important currency symbol. The rest can be ignored.

            Regex regex = new Regex(@"(?<s1>\\\s+)?\[\$(?<symbol>.+)-[\d\w]+\](?<s2>\\\s+)?");
            Match match = regex.Match(currencyCode);            

            if (match.Groups["symbol"].Success)
            {
                StringBuilder currencySymbol = new StringBuilder("");

                if (match.Groups["s1"].Success)
                    currencySymbol.Append(match.Groups["s1"].Value.Replace(@"\", ""));
                currencySymbol.Append(match.Groups["symbol"].Value);
                if (match.Groups["s2"].Success)
                    currencySymbol.Append(match.Groups["s2"].Value.Replace(@"\", ""));

                return currencySymbol.ToString();
            }
            else
                return string.Empty;
        }

        private bool TrySetFractionFormat(string formatCode)
        {
            // fractions not supported at the moment. Really - when are they used?!

            return false;

            //Regex fractionRegex = new Regex(@"^# \?+/\?+%");
            //Match fractionMatch = fractionRegex.Match(formatCode);

            //if (fractionMatch.Success)
            //{
            //    // There is no equivalent .NET format string to convert a decimal to a fraction. 
            //    // Need to develop method in this class to convert a decimal to a fraction

            //    // Leave as General for the time being

            //    return true;
            //}
        }

        private bool TrySetDateTimeFormat(string formatCode)
        {
            // [h], [m], [s] increases the time component beyond the maximum. Only one can be used in a format string. 
            // So 1.5 = 36:00:00 if format = [h]:mm:ss
            // Don't support AM/PM through complications with using locale codes (see below).
            // The code at the start of some default datetime formats (e.g. [$-F400] signifes locale codes. 
            // These indicate the system defaults should be used. The full potential format code is saved in the XML, and then ties in with
            // the system default.
            Regex dateTimeRegex = new Regex(@"^(?<c>\[.+\])?((?<y>y+)|(?<m>m+)|(?<d>d+)|(?<h>h+)|(?<s>s+)|(?<t>AM/PM)|(?<l>.))+$");
            Match dateTimeMatch = dateTimeRegex.Match(formatCode);

            if (dateTimeMatch.Groups["y"].Success || dateTimeMatch.Groups["m"].Success || dateTimeMatch.Groups["d"].Success ||
                dateTimeMatch.Groups["h"].Success || dateTimeMatch.Groups["s"].Success)
            {
                StringBuilder dotNetFormatCode = new StringBuilder("");

                FormatType = NumberFormatType.DateTime;

                // if there is only an m, then it is just month format.
                if (new Regex("^m+$").Match(formatCode).Success)
                {
                    DotNetFormatCode = formatCode.ToUpper();
                    return true;
                }

                // remove excel date time format code
                if (dateTimeMatch.Groups["c"].Success)
                    formatCode = formatCode.Replace(dateTimeMatch.Groups["c"].Value, "");
                // replace AM/PM code with dot net code
                //if (dateTimeMatch.Groups["t"].Success)
                //    formatCode = formatCode.Replace(dateTimeMatch.Groups["t"].Value, "tt");

                //// replace month part with upper case M for dot net format. Loop through each contiguous (non-whitespace) capture.                
                foreach (string formatPart in formatCode.Split(' '))
                {
                    bool isDate = formatPart.Contains("y") || formatPart.Contains("d");
                    bool isTime = formatPart.Contains("h") || formatPart.Contains("s");
                    bool isM = new Regex(@"^\W*m+\W*$").Match(formatPart.ToLower()).Success;

                    if (isDate || isTime || isM)
                    {
                        if (dotNetFormatCode.Length > 0)
                            dotNetFormatCode.Append(" ");

                        if (isDate)
                            dotNetFormatCode.Append(new Regex("[yMd].*[yMd]").Match(formatPart.Replace("m", "M")).Value);

                        if (isTime)
                            dotNetFormatCode.Append((new Regex("[Hms].*[Hms]").Match(formatPart.Replace("h", "H")).Value));

                        if (isM)
                            dotNetFormatCode.Append((new Regex("M+").Match(formatPart.ToUpper()).Value));
                    }
                }

                DotNetFormatCode = dotNetFormatCode.ToString();

                return true;
            }

            return false;
        }

        /***********************************
         * DAL METHODS
         ************************************/
        
        // Read

        internal static NumberFormat ReadNumberFormatFromReader(CustomOpenXmlReader reader, Styles styles, int numberFormatId)
        {
            NumberFormat numberFormat = new NumberFormat(styles);
            
            foreach (CustomOpenXmlAttribute attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "numFmtId":
                        numberFormat.NumberFormatId = attribute.GetIntValue();
                        break;
                    case "formatCode":
                        numberFormat.FormatCode = attribute.Value;
                        break;                    
                }
            }

            return numberFormat;
        }

        // Write

        internal static void WriteNumberFormatToWriter(CustomOpenXmlWriter<OpenXmlPackaging.WorkbookStylesPart> writer, NumberFormat numberFormat)
        {
            if (numberFormat.NumberFormatId != NumberFormat.DefaultNumberFormatId)
            {
                writer.WriteOpenXmlElement(new OpenXmlSpreadsheet.NumberingFormat());

                writer.WriteAttribute("numFmtId", numberFormat.NumberFormatId);
                writer.WriteAttribute("formatCode", numberFormat.FormatCode);

                writer.WriteEndElement();   // NumberingFormat
            }
        }

    }
}
