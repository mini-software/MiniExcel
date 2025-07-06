using MiniExcelLib.Core.OpenXml.Attributes;

namespace MiniExcelLib.Core.OpenXml.Utils;

internal static class ExcelPropertyHelper
{
    internal static ExcellSheetInfo GetExcellSheetInfo(Type type, MiniExcelBaseConfiguration configuration)
    {
        // default options
        var sheetInfo = new ExcellSheetInfo
        {
            Key = type.Name,
            ExcelSheetName = null, // will be generated automatically as Sheet<Index>
            ExcelSheetState = SheetState.Visible
        };

        // options from ExcelSheetAttribute
        if (type.GetCustomAttribute(typeof(ExcelSheetAttribute)) is ExcelSheetAttribute excelSheetAttr)
        {
            sheetInfo.ExcelSheetName = excelSheetAttr.Name ?? type.Name;
            sheetInfo.ExcelSheetState = excelSheetAttr.State;
        }

        // options from DynamicSheets configuration
        var openXmlCOnfiguration = configuration as OpenXmlConfiguration;
        if (openXmlCOnfiguration?.DynamicSheets?.Length > 0)
        {
            var dynamicSheet = openXmlCOnfiguration.DynamicSheets.SingleOrDefault(x => x.Key == type.Name);
            if (dynamicSheet is not null)
            {
                sheetInfo.ExcelSheetName = dynamicSheet.Name;
                sheetInfo.ExcelSheetState = dynamicSheet.State;
            }
        }

        return sheetInfo;
    }
}