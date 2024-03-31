namespace MiniExcelLibs.Attributes
{
    using System;
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        public string ExcelColumnName { get; set; }
        public string[] Aliases { get; set; }
        public ExcelColumnNameAttribute(string excelColumnName, string[] aliases = null)
        {
            ExcelColumnName = excelColumnName;
            Aliases = aliases;
        }
    }
}
