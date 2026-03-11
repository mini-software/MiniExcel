using System;

namespace MiniExcelLibs.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ExcelHiddenAttribute : Attribute
    {
        public bool ExcelHidden { get; set; }
        public ExcelHiddenAttribute(bool excelHidden = true) => ExcelHidden = excelHidden;
    }
}
