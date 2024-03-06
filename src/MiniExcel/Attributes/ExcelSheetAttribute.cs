using MiniExcelLibs.OpenXml;
using System;

namespace MiniExcelLibs.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ExcelSheetAttribute : Attribute
    {
        public string Name { get; set; }
        public SheetState State { get; set; } = SheetState.Visible;
    }

    public class DynamicExcelSheet : ExcelSheetAttribute
    {
        public string Key { get; set; }
        public DynamicExcelSheet(string key)
        {
            Key = key;
        }
    }
}