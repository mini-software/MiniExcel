namespace MiniExcelLibs
{
    using MiniExcelLibs.Csv;
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.OpenXml.SaveByTemplate;
    using System;
    using System.IO;

    internal static class ExcelReaderFactory
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
                    throw new NotSupportedException("Something went wrong. Please report this issue you are experiencing with MiniExcel.");
            }
        }
    }

    internal static class ExcelWriterFactory
    {
        internal static IExcelWriter GetProvider(Stream stream, object value, string sheetName, ExcelType excelType, IConfiguration configuration, bool printHeader)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("Sheet names can not be empty or null", nameof(sheetName));
            if (sheetName.Length > 31 && excelType == ExcelType.XLSX)
                throw new ArgumentException("Sheet names must be less than 31 characters", nameof(sheetName));
            if (excelType == ExcelType.UNKNOWN)
                throw new ArgumentException("Excel type cannot be ExcelType.UNKNOWN", nameof(excelType));

            switch (excelType)
            {
                case ExcelType.CSV:
                    return new CsvWriter(stream, value, configuration, printHeader);
                case ExcelType.XLSX:
                    return new ExcelOpenXmlSheetWriter(stream, value, sheetName, configuration, printHeader);
                default:
                    throw new NotSupportedException($"The {excelType} Excel format is not supported");
            }
        }
    }

    internal static class ExcelTemplateFactory
    {
        internal static IExcelTemplateAsync GetProvider(Stream stream, IConfiguration configuration, ExcelType excelType = ExcelType.XLSX)
        {
            switch (excelType)
            {
                case ExcelType.XLSX:
                    var valueExtractor = new InputValueExtractor();
                    return new ExcelOpenXmlTemplate(stream, configuration, valueExtractor);
                default:
                    throw new NotSupportedException("Something went wrong. Please report this issue you are experiencing with MiniExcel.");
            }
        }
    }
}
