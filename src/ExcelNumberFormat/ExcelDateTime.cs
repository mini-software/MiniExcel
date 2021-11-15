using System;
using System.Globalization;

namespace ExcelNumberFormat
{
    /// <summary>
    /// Similar to regular .NET DateTime, but also supports 0/1 1900 and 29/2 1900.
    /// </summary>
    internal class ExcelDateTime
    {
        /// <summary>
        /// The closest .NET DateTime to the specified excel date. 
        /// </summary>
        public DateTime AdjustedDateTime { get; }

        /// <summary>
        /// Number of days to adjust by in post.
        /// </summary>
        public int AdjustDaysPost { get; }

        /// <summary>
        /// Constructs a new ExcelDateTime from a numeric value.
        /// </summary>
        public ExcelDateTime(double numericDate, bool isDate1904)
        {
            if (isDate1904)
            {
                numericDate += 1462.0;
                AdjustedDateTime = new DateTime(DoubleDateToTicks(numericDate), DateTimeKind.Unspecified);
            }
            else
            {
                // internal dates before 30/12/1899 should add two days to get the real date
                // internal dates on 30/12 19899 should add two days, but subtract a day post to get the real date
                // internal dates before 28/2/1900 should add one day to get the real date
                // internal dates on 28/2 1900 should use the same date, but add a day post to get the real date

                var internalDateTime = new DateTime(DoubleDateToTicks(numericDate), DateTimeKind.Unspecified);
                if (internalDateTime < Excel1900ZeroethMinDate)
                {
                    AdjustDaysPost = 0;
                    AdjustedDateTime = internalDateTime.AddDays(2);
                }

                else if (internalDateTime < Excel1900ZeroethMaxDate)
                {
                    AdjustDaysPost = -1;
                    AdjustedDateTime = internalDateTime.AddDays(2);
                }

                else if (internalDateTime < Excel1900LeapMinDate)
                {
                    AdjustDaysPost = 0;
                    AdjustedDateTime = internalDateTime.AddDays(1);
                }

                else if (internalDateTime < Excel1900LeapMaxDate)
                {
                    AdjustDaysPost = 1;
                    AdjustedDateTime = internalDateTime;
                }
                else
                {
                    AdjustDaysPost = 0;
                    AdjustedDateTime = internalDateTime;
                }
            }
        }

        static DateTime Excel1900LeapMinDate = new DateTime(1900, 2, 28);
        static DateTime Excel1900LeapMaxDate = new DateTime(1900, 3, 1);
        static DateTime Excel1900ZeroethMinDate = new DateTime(1899, 12, 30);
        static DateTime Excel1900ZeroethMaxDate = new DateTime(1899, 12, 31);

        /// <summary>
        /// Wraps a regular .NET datetime.
        /// </summary>
        /// <param name="value"></param>
        public ExcelDateTime(DateTime value)
        {
            AdjustedDateTime = value;
            AdjustDaysPost = 0;
        }

        public int Year => AdjustedDateTime.Year;

        public int Month => AdjustedDateTime.Month;

        public int Day => AdjustedDateTime.Day + AdjustDaysPost;

        public int Hour => AdjustedDateTime.Hour;

        public int Minute => AdjustedDateTime.Minute;

        public int Second => AdjustedDateTime.Second;

        public int Millisecond => AdjustedDateTime.Millisecond;

        public DayOfWeek DayOfWeek => AdjustedDateTime.DayOfWeek;

        public string ToString(string numberFormat, CultureInfo culture)
        {
            return AdjustedDateTime.ToString(numberFormat, culture);
        }

        public static bool TryConvert(object value, bool isDate1904, CultureInfo culture, out ExcelDateTime result)
        {
            if (value is double doubleValue)
            {
                result = new ExcelDateTime(doubleValue, isDate1904);
                return true;
            }
            if (value is int intValue)
            {
                result = new ExcelDateTime(intValue, isDate1904);
                return true;
            }
            if (value is short shortValue)
            {
                result = new ExcelDateTime(shortValue, isDate1904);
                return true;
            }
            else if (value is DateTime dateTimeValue)
            {
                result = new ExcelDateTime(dateTimeValue);
                return true;
            }

            result = null;
            return false;
        }

        // From DateTime class to enable OADate in PCL
        // Number of 100ns ticks per time unit
        private const long TicksPerMillisecond = 10000;
        private const long TicksPerSecond = TicksPerMillisecond * 1000;
        private const long TicksPerMinute = TicksPerSecond * 60;
        private const long TicksPerHour = TicksPerMinute * 60;
        private const long TicksPerDay = TicksPerHour * 24;

        private const int MillisPerSecond = 1000;
        private const int MillisPerMinute = MillisPerSecond * 60;
        private const int MillisPerHour = MillisPerMinute * 60;
        private const int MillisPerDay = MillisPerHour * 24;

        // Number of days in a non-leap year
        private const int DaysPerYear = 365;

        // Number of days in 4 years
        private const int DaysPer4Years = DaysPerYear * 4 + 1;

        // Number of days in 100 years
        private const int DaysPer100Years = DaysPer4Years * 25 - 1;

        // Number of days in 400 years
        private const int DaysPer400Years = DaysPer100Years * 4 + 1;

        // Number of days from 1/1/0001 to 12/30/1899
        private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

        private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;

        internal static long DoubleDateToTicks(double value)
        {
            long millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

            // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
            // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
            // This line below fixes up the millis in the negative case
            if (millis < 0)
            {
                millis -= millis % MillisPerDay * 2;
            }

            millis += DoubleDateOffset / TicksPerMillisecond;
            return millis * TicksPerMillisecond;
        }
    }
}
