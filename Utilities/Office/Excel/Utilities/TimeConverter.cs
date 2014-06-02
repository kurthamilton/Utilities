using System;

namespace Utilities.Office.Excel
{
    internal class TimeConverter : BaseExcel
    {
        // In practice the 0th date in Excel is 00/01/1900 = 31/12/1899. 
        // Unfortunately (and by design to be compatible with Lotus 123) Excel considers 29/02/1900 to exist, 
        // while SQL Server and Visual Studio don't because it didn't really happen. 
        // To account for this, it's easier to pretend the Excel base date is 30/12/1899, and account for this when handling dates <= 01/03/1900
        private DateTime _baseExcelDateTime;
        protected DateTime BaseExcelDateTime { get { if (_baseExcelDateTime == DateTime.MinValue) _baseExcelDateTime = new DateTime(1899, 12, 30); return _baseExcelDateTime; } }

        private DateTime _additionalExcelDate;
        protected DateTime AdditionalExcelDate { get { if (_additionalExcelDate == DateTime.MinValue) _additionalExcelDate = new DateTime(1900, 3, 1); return _additionalExcelDate; } } 

        private DateTime _baseExternalDateTime;
        protected DateTime BaseExternalDateTime { get { if (_baseExternalDateTime == DateTime.MinValue) _baseExternalDateTime = new DateTime(1900, 1, 1); return _baseExternalDateTime; } }

        private TimeSpan _baseTimeSpan;
        protected TimeSpan BaseTimeSpan { get { if (_baseTimeSpan == TimeSpan.Zero) _baseTimeSpan = BaseExternalDateTime.Subtract(BaseExcelDateTime); return _baseTimeSpan; } }

        double hourRatio;
        double minuteRatio;
        double secondRatio;
        double millisecondRatio;

        public TimeConverter()
        {
            hourRatio = (double)1 / (double)24;
            minuteRatio = hourRatio / (double)60;
            secondRatio = minuteRatio / (double)60;
            millisecondRatio = secondRatio / (double)1000;
        }

        public double GetDoubleValueFromDateTimeValue(DateTime value)
        {
            if (GetDateFromDateTime(value).CompareTo(BaseExternalDateTime) == 0)
                value = GetExcelTimeFromDateTime(value);
            else
                value = GetCorrectedDateForExcel(value);

            TimeSpan timeSpan = value.Subtract(BaseExternalDateTime).Add(BaseTimeSpan);
            return GetDoubleTimeSpanFromTimeSpan(timeSpan);
        }

        public DateTime GetDateTimeValueFromDoubleValue(double value)
        {
            if (value < 0)
                throw new InvalidCastException();

            try
            {
                int days = (int)Math.Floor(value);

                DateTime dateTimeValue = BaseExcelDateTime;
                dateTimeValue = dateTimeValue.AddDays((double)days);

                double timePart = value - days;

                int hours = (int)Math.Floor(timePart / hourRatio);
                timePart = timePart - ((double)hours * hourRatio);
                int minutes = (int)Math.Floor(timePart / minuteRatio);
                timePart = timePart - ((double)minutes * minuteRatio);
                int seconds = (int)Math.Round(timePart / secondRatio, 0);

                dateTimeValue = dateTimeValue.Add(new TimeSpan(hours, minutes, seconds));

                return GetCorrectedDateFromExcel(dateTimeValue);
            }
            catch (ArgumentOutOfRangeException)
            {
                // catch values that are too large for DateTime.                
            }

            return DateTime.MinValue;
        }

        protected double GetDoubleTimeSpanFromTimeSpan(TimeSpan timeSpan)
        {
            double doubleValue = 0;
            doubleValue += timeSpan.Days;

            doubleValue += (double)timeSpan.Hours * hourRatio;
            doubleValue += (double)timeSpan.Minutes * minuteRatio;
            doubleValue += (double)timeSpan.Seconds * secondRatio;
            // Excel rounds up seconds when milliseconds are included, as far as I can see, so ignore
            //doubleValue += (double)timeSpan.Milliseconds * millisecondRatio;

            return doubleValue;
        }

        protected DateTime GetCorrectedDateForExcel(DateTime value)
        {
            // see note for BaseExcelDateTime class property for explanation

            if (value.CompareTo(AdditionalExcelDate) < 0)
                return value.Subtract(new TimeSpan(1, 0, 0, 0));
            else
                return value;
        }

        protected DateTime GetCorrectedDateFromExcel(DateTime value)
        {
            // see note for BaseExcelDateTime class property for explanation

            if (value.CompareTo(AdditionalExcelDate) < 0)
                return value.Add(new TimeSpan(1, 0, 0, 0));
            else
                return value;
        }

        protected DateTime GetDateFromDateTime(DateTime value)
        {
            DateTime dateValue = value.Subtract(GetTimeTimeSpanFromDateTime(value));
            return dateValue;
        }

        protected DateTime GetExcelTimeFromDateTime(DateTime value)
        {
            DateTime timeValue = BaseExcelDateTime.Add(GetTimeTimeSpanFromDateTime(value));
            return timeValue;
        }

        protected TimeSpan GetTimeTimeSpanFromDateTime(DateTime value)
        {
            TimeSpan timeTimeSpan = new TimeSpan(0, value.Hour, value.Minute, value.Second, value.Millisecond);
            return timeTimeSpan;
        }
        
    }
}
