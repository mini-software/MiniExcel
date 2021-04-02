namespace MiniExcelLibs.Utils
{
    using System;

    internal static partial class Helpers
    {
        [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
        public class ExcelColumnNameAttribute : Attribute
        {
            public string ExcelColumnName { get; set; }
            public ExcelColumnNameAttribute(string excelColumnName) => ExcelColumnName = excelColumnName;
        }

    }

}
