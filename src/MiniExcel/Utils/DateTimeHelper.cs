namespace MiniExcelLibs.Utils
{
    public static class DateTimeHelper
    {
        /// <summary>
        /// NumberFormat from NuGet ExcelNumberFormat MIT@License
        /// </summary>
        public static bool IsDateTimeFormat(string formatCode)
        {
            return new ExcelNumberFormat(formatCode).IsDateTimeFormat;
        }

        /**Below Code from ExcelDataReader @MIT License**/
        // All OA dates must be strictly in between OADateMinAsDouble and OADateMaxAsDouble
        private const double OADateMinAsDouble = -657435.0;

        private const double OADateMaxAsDouble = 2958466.0;

        internal static bool IsValidOADateTime(double value)
        {
            return value > OADateMinAsDouble && value < OADateMaxAsDouble;
        }
    }
}
