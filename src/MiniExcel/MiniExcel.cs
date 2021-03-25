namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System.Linq;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System;
    using MiniExcelLibs.Csv;
    public static partial class MiniExcel
    {
        public static void SaveAs(this Stream stream, object value, string startCell = "A1", bool printHeader = true, ExcelType excelType = ExcelType.XLSX)
        {
            switch (excelType)
            {
                case ExcelType.CSV:
                    CsvWriter.SaveAs(stream, value);
                    break;
                case ExcelType.XLSX:
                    SaveAsImpl(stream, GetCreateXlsxInfos(value, startCell, printHeader));
                    break;
                default:
                    throw new NotSupportedException($"Extension : {excelType} not suppprt");
            }
        }

        public static void SaveAs(string filePath, object value, string startCell = "A1", bool printHeader = true, ExcelType excelType = ExcelType.UNKNOWN)
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(filePath);
            switch (excelType)
            {
                case ExcelType.CSV:
                    CsvWriter.SaveAs(filePath, value);
                    break;
                case ExcelType.XLSX:
                    SaveAsImpl(filePath, GetCreateXlsxInfos(value, startCell, printHeader));
                    break;
                default:
                    throw new NotSupportedException($"Extension : {Path.GetExtension(filePath)} not suppprt");
            }
        }

        public static IEnumerable<T> Query<T>(string path) where T : class, new()
        {
            using (var stream = File.OpenRead(path))
            {
                return QueryImpl<T>(stream);
            }
        }

        public static IEnumerable<T> Query<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream);
        }

        public static T QueryFirst<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).First();
        }

        public static T QueryFirstOrDefault<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).FirstOrDefault();
        }

        public static T QuerySingle<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).Single();
        }

        public static T QuerySingleOrDefault<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).SingleOrDefault();
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, ExcelType excelType = ExcelType.UNKNOWN,IConfiguration configuration=null) 
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(path);
            //using (var stream = File.OpenRead(path))
            Stream stream = null;
            {
                switch (excelType)
                {
                    case ExcelType.CSV:
                        return new CsvReader().Query(path, useHeaderRow, (CsvConfiguration)configuration);
                    case ExcelType.XLSX:
                        return new ExcelOpenXmlSheetReader().Query(stream, useHeaderRow);
                    default:
                        throw new NotSupportedException($"Extension : {Path.GetExtension(path)} not suppprt");
                }
            }
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            if (excelType == ExcelType.UNKNOWN)
                excelType = GetExcelType(stream);
            switch (excelType)
            {
                case ExcelType.CSV:
                    return new CsvReader().Query(stream, useHeaderRow, (CsvConfiguration)configuration);
                case ExcelType.XLSX:
                    return new ExcelOpenXmlSheetReader().Query(stream, useHeaderRow);
                default:
                    throw new NotSupportedException($"Please Issue for me");
            }
        }

        public static dynamic QueryFirst(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().Query(stream, useHeaderRow).First();
        }

        public static dynamic QueryFirstOrDefault(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().Query(stream, useHeaderRow).FirstOrDefault();
        }

        public static dynamic QuerySingle(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().Query(stream, useHeaderRow).Single();
        }

        public static dynamic QuerySingleOrDefault(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().Query(stream, useHeaderRow).SingleOrDefault();
        }
    }
}
