namespace MiniExcelLibs.OpenXml
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Parse ECMA-376 number format strings from Excel and other spreadsheet softwares.
    /// </summary>
    public class FormatTypeMapping
    {
        private static Dictionary<int, FormatTypeMapping> Formats { get; } = new Dictionary<int, FormatTypeMapping>()
        {
            { 0, new FormatTypeMapping("General",typeof(string)) },
            { 1, new FormatTypeMapping("0",typeof(int?)) },
            { 2, new FormatTypeMapping("0.00",typeof(double?)) },
            { 3, new FormatTypeMapping("#,##0",typeof(decimal?)) },
            { 4, new FormatTypeMapping("#,##0.00",typeof(decimal?)) },
            { 5, new FormatTypeMapping("\"$\"#,##0_);(\"$\"#,##0)") },
            { 6, new FormatTypeMapping("\"$\"#,##0_);[Red](\"$\"#,##0)") },
            { 7, new FormatTypeMapping("\"$\"#,##0.00_);(\"$\"#,##0.00)") },
            { 8, new FormatTypeMapping("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)") },
            { 9, new FormatTypeMapping("0%") },
            { 10, new FormatTypeMapping("0.00%") },
            { 11, new FormatTypeMapping("0.00E+00") },
            { 12, new FormatTypeMapping("# ?/?") },
            { 13, new FormatTypeMapping("# ??/??") },
            { 14, new FormatTypeMapping("d/m/yyyy",typeof(DateTime?)) },
            { 15, new FormatTypeMapping("d-mmm-yy",typeof(DateTime?)) },
            { 16, new FormatTypeMapping("d-mmm",typeof(DateTime?)) },
            { 17, new FormatTypeMapping("mmm-yy",typeof(DateTime?)) },
            { 18, new FormatTypeMapping("h:mm AM/PM",typeof(DateTime?)) },
            { 19, new FormatTypeMapping("h:mm:ss AM/PM",typeof(DateTime?)) },
            { 20, new FormatTypeMapping("h:mm",typeof(DateTime?)) },
            { 21, new FormatTypeMapping("h:mm:ss",typeof(DateTime?)) },
            { 22, new FormatTypeMapping("m/d/yy h:mm",typeof(DateTime?)) },

            // 23..36 international/unused
            { 37, new FormatTypeMapping("#,##0_);(#,##0)") },
            { 38, new FormatTypeMapping("#,##0_);[Red](#,##0)") },
            { 39, new FormatTypeMapping("#,##0.00_);(#,##0.00)") },
            { 40, new FormatTypeMapping("#,##0.00_);[Red](#,##0.00)") },
            { 41, new FormatTypeMapping("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)") },
            { 42, new FormatTypeMapping("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)") },
            { 43, new FormatTypeMapping("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)") },
            { 44, new FormatTypeMapping("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)") },
            { 45, new FormatTypeMapping("mm:ss",typeof(TimeSpan?)) },
            { 46, new FormatTypeMapping("[h]:mm:ss",typeof(TimeSpan?)) },
            { 47, new FormatTypeMapping("mm:ss.0",typeof(TimeSpan?)) },
            { 48, new FormatTypeMapping("##0.0E+0") },
            { 49, new FormatTypeMapping("@") },
        };

        public static FormatTypeMapping GetBuiltinNumberFormat(int numFmtId)
        {
            if (Formats.TryGetValue(numFmtId, out var result))
                return result;

            return null;
        }

        public FormatTypeMapping(string formatString, Type formatType = null)
        {
            if (formatType == null)
                FormatType = typeof(string);
            //TODO:Custom Date Check
            FormatString = formatString;
            FormatType = formatType;
        }
        public string FormatString { get; }

        public Type FormatType { get; }
    }
    public class ExtendedFormat
    {
        /// <summary>
        /// Gets or sets the index to the parent Cell Style CF record with overrides for this XF. Only used with Cell XFs.
        /// 0xFFF means no override
        /// </summary>
        public int ParentCellStyleXf { get; set; }
        public int NumberFormatIndex { get; set; }
    }

}
