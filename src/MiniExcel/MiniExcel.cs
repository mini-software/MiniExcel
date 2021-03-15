namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System.Linq;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.IO.Compression;
    using System.Text;
    using System.Reflection;
    using MiniExcelLibs.Utils;

    public static partial class MiniExcel
    {
        private static Dictionary<string, ZipPackageInfo> GetDefaultFiles() => new Dictionary<string, ZipPackageInfo>()
        {
            { @"_rels/.rels",new ZipPackageInfo(DefualtXml.DefaultRels, "application/vnd.openxmlformats-package.relationships+xml")},
            { @"xl/_rels/workbook.xml.rels",new ZipPackageInfo(DefualtXml.DefaultWorkbookXmlRels, "application/vnd.openxmlformats-package.relationships+xml")},
            { @"xl/styles.xml",new ZipPackageInfo(DefualtXml.DefaultStylesXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")},
            { @"xl/workbook.xml",new ZipPackageInfo(DefualtXml.DefaultWorkbookXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")},
            { @"xl/worksheets/sheet1.xml",new ZipPackageInfo(DefualtXml.DefaultSheetXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")},
        };

        private readonly static UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);

        public static void SaveAs(this Stream stream,object value, string startCell = "A1", bool printHeader = true)
        {
            SaveAsImpl(stream,GetCreateXlsxInfos(value, startCell, printHeader));
            stream.Position = 0;
        }

        public static void SaveAs(string filePath, object value, string startCell = "A1", bool printHeader = true)
        {
            CreateXlsxFile(filePath, GetCreateXlsxInfos(value, startCell, printHeader));
        }

        private static Dictionary<string, ZipPackageInfo> GetCreateXlsxInfos(object value, string startCell, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY(startCell);

            var defaultFiles = GetDefaultFiles();
            {
                var sb = new StringBuilder();

                var yIndex = xy.Item2;

                if (value is DataTable)
                {
                    var dt = value as DataTable;
                    if (printHeader)
                    {
                        sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (DataColumn c in dt.Columns)
                        {
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            sb.Append($"<x:c r=\"{columname}\" t=\"str\">");
                            sb.Append($"<x:v>{c.ColumnName}");
                            sb.Append($"</x:v>");
                            sb.Append($"</x:c>");
                            xIndex++;
                        }
                        sb.AppendLine($"</x:row>");
                        yIndex++;
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            var cellValue = dt.Rows[i][j];
                            var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                            var t = "t=\"str\"";
                            {
                                if (decimal.TryParse(cellValueStr, out var outV))
                                    t = "t=\"n\"";
                                if (cellValue is bool)
                                {
                                    t = "t=\"b\"";
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                if (cellValue is DateTime || cellValue is DateTime?)
                                {
                                    t = "s=\"1\"";
                                    cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                                }
                            }
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            sb.Append($"<x:c r=\"{columname}\" {t}>");
                            sb.Append($"<x:v>{cellValueStr}");
                            sb.Append($"</x:v>");
                            sb.Append($"</x:c>");
                            xIndex++;
                        }
                        sb.AppendLine($"</x:row>");
                        yIndex++;
                    }
                }
                else if (value is System.Collections.ICollection)
                {
                    var collection = value as System.Collections.ICollection;
                    object firstValue = null;
                    {
                        foreach (var v in collection)
                        {
                            firstValue = v;
                            break;
                        }
                    }
                    var type = firstValue.GetType();
                    var props = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);
                    if (printHeader)
                    {
                        sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (var p in props)
                        {
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            sb.Append($"<x:c r=\"{columname}\" t=\"str\">");
                            sb.Append($"<x:v>{p.Name}");
                            sb.Append($"</x:v>");
                            sb.Append($"</x:c>");
                            xIndex++;
                        }
                        sb.AppendLine($"</x:row>");
                        yIndex++;
                    }

                    foreach (var v in collection)
                    {
                        sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (var p in props)
                        {
                            var cellValue = p.GetValue(v);
                            var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                            var t = "t=\"str\"";
                            {
                                if (decimal.TryParse(cellValueStr, out var outV))
                                    t = "t=\"n\"";
                                if (cellValue is bool)
                                {
                                    t = "t=\"b\"";
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                if (cellValue is DateTime || cellValue is DateTime?)
                                {
                                    t = "s=\"1\"";
                                    cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                                }
                            }
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            sb.Append($"<x:c r=\"{columname}\" {t}>");
                            sb.Append($"<x:v>{cellValueStr}");
                            sb.Append($"</x:v>");
                            sb.Append($"</x:c>");
                            xIndex++;
                        }
                        sb.AppendLine($"</x:row>");
                        yIndex++;
                    }
                }

                defaultFiles[@"xl/worksheets/sheet1.xml"].Xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<x:sheetData>{sb.ToString()}</x:sheetData>
</x:worksheet>";
            }

            return defaultFiles;
        }

        public static IEnumerable<T> Query<T>(this Stream stream) where T : class , new()
        {
            return QueryImpl<T>(stream);
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow);
        }

        public static dynamic QueryFirst(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).First();
        }

        public static dynamic QueryFirstOrDefault(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).FirstOrDefault();
        }

        public static dynamic QuerySingle(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).Single();
        }

        public static dynamic QuerySingleOrDefault(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).SingleOrDefault();
        }

        private static IEnumerable<T> QueryImpl<T>(this Stream stream) where T : class, new()
        {
            var type = typeof(T);
            var props = Helpers.GetPropertiesWithSetter(type);
            foreach (var item in new ExcelOpenXmlSheetReader().QueryImpl(stream, true))
            {
                var v = new T();
                foreach (var p in props)
                {
                    if (item.ContainsKey(p.Name))
                    {
                        object newV = null;
                        if (p.PropertyType == typeof(Guid))
                            newV = Guid.Parse(item[p.Name].ToString());
                        else if (p.PropertyType == typeof(DateTime))
                        {
                            var bs = item[p.Name].ToString();
                            if (DateTime.TryParse(bs, out var _v))
                                newV = _v;
                            else
                                newV = DateTime.ParseExact(bs, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else
                            newV = Convert.ChangeType(item[p.Name], p.PropertyType);
                        p.SetValue(v, newV);
                    }
                }
                yield return v;
            }
        }



        private static void CreateXlsxFile(string path, Dictionary<string, ZipPackageInfo> zipPackageInfos)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            using(ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false, Utf8WithBom))
                CreteXlsxImpl(zipPackageInfos, archive);
        }
        private static void SaveAsImpl(Stream stream,Dictionary<string, ZipPackageInfo> zipPackageInfos)
        {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                CreteXlsxImpl(zipPackageInfos, archive);
            }
        }

        private static void CreteXlsxImpl(Dictionary<string, ZipPackageInfo> zipPackageInfos, ZipArchive archive)
        {
            //[Content_Types].xml
            {
                var sb = new StringBuilder(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default ContentType=""application/xml"" Extension=""xml""/>
    <Default ContentType=""application/vnd.openxmlformats-package.relationships+xml"" Extension=""rels""/>");
                foreach (var p in zipPackageInfos)
                {
                    sb.AppendLine($"<Override ContentType=\"{p.Value.ContentType}\" PartName=\"/{p.Key}\" />");
                }
                sb.AppendLine("</Types>");

                ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml");
                using (var zipStream = entry.Open())
                using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    writer.Write(sb.ToString());
            }

            foreach (var p in zipPackageInfos)
            {
                ZipArchiveEntry entry = archive.CreateEntry(p.Key);
                using (var zipStream = entry.Open())
                using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    writer.Write(p.Value.Xml.ToString());
            }
        }
    }
}
