namespace MiniExcelLibs.Attributes
{
    using MiniExcelLibs.Utils;
    using System;
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        public string ExcelColumnName { get; set; }
        public ExcelColumnNameAttribute(string excelColumnName) => ExcelColumnName = excelColumnName;
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelColumnIndexAttribute : Attribute
    {
	   public int ExcelColumnIndex { get; set; }
	   public ExcelColumnIndexAttribute(string columnName) => ExcelColumnIndex = Helpers.GetColumnIndex(columnName);
	   public ExcelColumnIndexAttribute(int columnIndex) => ExcelColumnIndex = columnIndex;
    }
}
