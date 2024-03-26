namespace MiniExcelLibs
{
    using MiniExcelLibs.Csv;
    using MiniExcelLibs.OpenXml;
    using System;
    using System.IO;

    internal class ExcelReaderFactory
    {
        internal static IExcelReader GetProvider(Stream stream, ExcelType excelType, IConfiguration configuration)
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

    internal class ExcelWriterFactory
    {
        internal static IExcelWriter GetProvider(Stream stream, object value, string sheetName, ExcelType excelType, IConfiguration configuration, bool printHeader)
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
    }

    internal class ExcelTemplateFactory
    {
        internal static IExcelTemplateAsync GetProvider(Stream stream, IConfiguration configuration, ExcelType excelType = ExcelType.XLSX)
        {
            switch (excelType)
            {
                case ExcelType.XLSX:
                    return new ExcelOpenXmlTemplate(stream, configuration);
                default:
                    throw new NotSupportedException($"Please Issue for me");
            }
        }
    }
}
