using MiniExcelLib.Core.Attributes;

// ReSharper disable CheckNamespace
namespace MiniExcelLibs.Attributes;


[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelColumnAttribute : MiniExcelColumnAttribute;

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelColumnIndexAttribute : MiniExcelColumnIndexAttribute
{
    public ExcelColumnIndexAttribute(int excelColumnIndex) : base(excelColumnIndex) { }
    public ExcelColumnIndexAttribute(string excelColumnName) : base(excelColumnName) { }
}

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelColumnNameAttribute(string excelColumnName) : MiniExcelColumnNameAttribute(columnName: excelColumnName);

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelColumnWidthAttribute(double width) : MiniExcelColumnWidthAttribute(width);

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelFormatAttribute(string format) : MiniExcelFormatAttribute(format);

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelIgnoreAttribute(bool excelIgnore = true) : MiniExcelIgnoreAttribute(excelIgnore);

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class ExcelSheetAttribute : MiniExcelSheetAttribute;

[Obsolete("This is a legacy attribute that will be removed in a future version. Please use the corresponding one from MiniExcelLib.Core.Attributes instead.")]
public sealed class DynamicExcelSheetAttribute(string key) : MiniExcelLib.Core.Attributes.DynamicExcelSheetAttribute(key);