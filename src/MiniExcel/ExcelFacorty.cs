namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System;
    using MiniExcelLibs.Csv;
    using System.IO;
    using System.Globalization;

    internal class ExcelReaderFactory
    { 
        internal static IExcelReader GetProvider(Stream stream, ExcelType excelType,IConfiguration configuration)
        {
            switch (excelType)
            {
                case ExcelType.CSV:
                    return new CsvReader(stream, configuration);
                case ExcelType.XLSX:
                    return new ExcelOpenXmlSheetReader(stream, configuration);
                default:
                    throw new NotSupportedException($"Please Issue for me");
            }
        }
    }

    internal class ExcelTemplateFactory
    {
        internal static IExcelTemplateAsync GetProvider(Stream stream, ExcelType excelType= ExcelType.XLSX)
        {
            switch (excelType)
            {
                case ExcelType.XLSX:
                    return new ExcelOpenXmlTemplate(stream);
                default:
                    throw new NotSupportedException($"Please Issue for me");
            }
        }
    }
}
