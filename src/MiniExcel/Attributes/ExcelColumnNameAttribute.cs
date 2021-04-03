namespace MiniExcelLibs.Attributes
{
    using System;
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        public string ExcelColumnName { get; set; }

        public int? ExcelColumnIndex { get; set; }

        public ExcelColumnNameAttribute(string excelColumnName) : this(excelColumnName,0) { }

        public ExcelColumnNameAttribute(string excelColumnName, int columnIndex) {
            ExcelColumnName = excelColumnName;
            ExcelColumnIndex = columnIndex;
        }
    }

    [AttributeUsage(AttributeTargets.Class , AllowMultiple = false)]
    public class ExcelAttribute : Attribute
    {
        
    }
}
