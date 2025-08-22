using MiniExcelLib.Core.Attributes;

// ReSharper disable CheckNamespace
namespace MiniExcelLibs.Attributes;


public sealed class ExcelColumnAttribute : MiniExcelColumnAttribute;

public sealed class ExcelColumnIndexAttribute : MiniExcelColumnIndexAttribute
{
    public ExcelColumnIndexAttribute(int excelColumnIndex) : base(excelColumnIndex) { }
    public ExcelColumnIndexAttribute(string excelColumnName) : base(excelColumnName) { }
}

public sealed class ExcelColumnNameAttribute(string excelColumnName) : MiniExcelColumnNameAttribute(columnName: excelColumnName);

public sealed class ExcelColumnWidthAttribute(double width) : MiniExcelColumnWidthAttribute(width);

public sealed class ExcelFormatAttribute(string format) : MiniExcelFormatAttribute(format);

public sealed class ExcelIgnoreAttribute(bool excelIgnore = true) : MiniExcelIgnoreAttribute(excelIgnore);

public sealed class ExcelSheetAttribute : MiniExcelSheetAttribute;

public sealed class DynamicExcelSheetAttribute(string key) : MiniExcelLib.Core.Attributes.DynamicExcelSheetAttribute(key);