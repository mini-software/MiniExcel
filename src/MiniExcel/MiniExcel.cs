namespace MiniExcelLibs
{
    using System.Collections.Generic;
    using System.IO;

    public static partial class MiniExcel
    {
        public static void SaveAs(string path, object value, bool printHeader = true, ExcelType excelType = ExcelType.UNKNOWN)
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(path);
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, excelType);
        }

        /// <summary>
        /// Default SaveAs Xlsx
        /// </summary>
        public static void SaveAs(this Stream stream, object value, bool printHeader = true, ExcelType excelType = ExcelType.XLSX)
        {
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");
            ExcelFacorty.GetExcelProvider(excelType, printHeader).SaveAs(stream, value);
        }

        public static IEnumerable<T> Query<T>(string path, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(path);
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
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(path);
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
