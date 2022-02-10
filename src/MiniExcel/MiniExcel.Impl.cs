namespace MiniExcelLibs
{
    using MiniExcelLibs.Csv;
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Utils;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    public static partial class MiniExcel
    {
        private static IExcelWriter GetWriterProvider(Stream stream, object value, string sheetName, ExcelType excelType, IConfiguration configuration, bool printHeader)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidDataException("Sheet name can not be empty or null");
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");

            switch (excelType)
            {
                case ExcelType.CSV:
                    return new CsvWriter(stream, value, configuration, printHeader);
                case ExcelType.XLSX:
                    return new ExcelOpenXmlSheetWriter(stream, value, sheetName, configuration, printHeader);
                default:
                    throw new NotSupportedException($"Please Issue for me");
            }
        }

        private static IExcelReader GetReaderProvider(Stream stream, ExcelType excelType)
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType),null);
        }
    }
}
