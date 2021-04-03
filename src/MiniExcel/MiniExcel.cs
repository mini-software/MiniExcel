namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System.Linq;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System;
    using MiniExcelLibs.Csv;
    using System.Data;
    using System.Collections;

    public static partial class MiniExcel
    {
        public static void SaveAs(this Stream stream, object value, bool printHeader = true, ExcelType excelType = ExcelType.XLSX)
        {
            ExcelFacorty.GetExcelProvider(excelType, printHeader).SaveAs(stream, value);
        }

        public static void SaveAs(string filePath, object value, bool printHeader = true, ExcelType excelType = ExcelType.UNKNOWN)
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(filePath);

            ExcelFacorty.GetExcelProvider(excelType, printHeader).SaveAs(filePath, value);
        }

        public static IEnumerable<T> Query<T>(string path, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query<T>(stream, excelType, configuration))
                    yield return item;
        }

        public static IEnumerable<T> Query<T>(this Stream stream, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(stream);

            return ExcelFacorty.GetExcelProvider(excelType).Query<T>(stream);
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) 
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query(stream, useHeaderRow, excelType, configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(stream);

            return ExcelFacorty.GetExcelProvider(excelType, useHeaderRow).Query(stream, useHeaderRow);
        }
    }
}
