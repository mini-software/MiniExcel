namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System;
    using MiniExcelLibs.Csv;

    /// <summary>
    /// use statics factory,If want to do OCP we can use a Interface factory to instead of  statics factory
    /// </summary>
    internal class ExcelFacorty {

        internal static ExcelProviderBase GetExcelProvider(ExcelType excelType) {
            return GetExcelProvider(excelType,true);
        }
        internal static ExcelProviderBase GetExcelProvider(ExcelType excelType,bool useHeaderRow) {
            switch (excelType)
            {
                case ExcelType.CSV:
                    return new CsvProvider();
                case ExcelType.XLSX:
                    return new ExcelOpenXmlProvider(useHeaderRow);
                default:
                    throw new NotSupportedException($"Please Issue for me");
            }
        }
    }
}
