namespace MiniExcelLibs.Utils
{
    using ExcelNumberFormat;
    using System;
    using System.Globalization;

    internal static partial class DateTimeHelper
    {
        /// <summary>
        /// NumberFormat from NuGet ExcelNumberFormat MIT@License
        /// </summary>
        public static bool IsDateTimeFormat(string formatCode)
        {
            return new NumberFormat(formatCode).IsDateTimeFormat;
        }
    }

    internal static partial class DateTimeHelper
    {
        /**Below Code from ExcelDataReader @MIT License**/

        // All OA dates must be greater than (not >=) OADateMinAsDouble
        public const double OADateMinAsDouble = -657435.0;

        // All OA dates must be less than (not <=) OADateMaxAsDouble
        public const double OADateMaxAsDouble = 2958466.0;

        // From DateTime class to enable OADate in PCL
        // Number of 100ns ticks per time unit
        private const long TicksPerMillisecond = 10000;
        private const long TicksPerSecond = TicksPerMillisecond * 1000;
        private const long TicksPerMinute = TicksPerSecond * 60;
        private const long TicksPerHour = TicksPerMinute * 60;
        private const long TicksPerDay = TicksPerHour * 24;

        // Number of milliseconds per time unit
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

        // Number of days from 1/1/0001 to 12/31/9999
        private const int DaysTo10000 = DaysPer400Years * 25 - 366;

        private const long MaxMillis = (long)DaysTo10000 * MillisPerDay;

        private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;

        public static DateTime FromOADate(double d)
        {
            return new DateTime(DoubleDateToTicks(d), DateTimeKind.Unspecified);
        }

        // duplicated from DateTime
        internal static long DoubleDateToTicks(double value)
        {
            if (value >= OADateMaxAsDouble || value <= OADateMinAsDouble)
                throw new ArgumentException("Invalid OA Date");
            long millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

            // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
            // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
            // This line below fixes up the millis in the negative case
            if (millis < 0)
            {
                millis -= millis % MillisPerDay * 2;
            }

            millis += DoubleDateOffset / TicksPerMillisecond;

            if (millis < 0 || millis >= MaxMillis)
                throw new ArgumentException("OA Date out of range");
            return millis * TicksPerMillisecond;
        }

        public static double AdjustOADateTime(double value, bool date1904)
        {
            if (!date1904)
            {
                // Workaround for 1900 leap year bug in Excel
                if (value >= 0.0 && value < 60.0)
                    return value + 1;
            }
            else
            {
                return value + 1462.0;
            }

            return value;
        }

        public static bool IsValidOADateTime(double value)
        {
            return value > DateTimeHelper.OADateMinAsDouble && value < DateTimeHelper.OADateMaxAsDouble;
        }

        public static object ConvertFromOATime(double value, bool date1904)
        {
            var dateValue = AdjustOADateTime(value, date1904);
            if (IsValidOADateTime(dateValue))
                return DateTimeHelper.FromOADate(dateValue);
            return value;
        }
    }

}
