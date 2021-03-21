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
    using System.Globalization;
    using System.Collections;

    public static partial class MiniExcel
    {
        private static Dictionary<string, ZipPackageInfo> GetCreateXlsxInfos(object value, string startCell, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY(startCell);

            var defaultFiles = DefualtXml.GetDefaultFiles();

            // dimension
            var dimensionRef = "A1";
            var maxRowIndex = 0;
            var maxColumnIndex = 0;

            {
                var sb = new StringBuilder();
                var yIndex = xy.Item2;

                if (value is DataTable)
                {
                    var dt = value as DataTable;

                    // dimension
                    maxRowIndex = dt.Rows.Count;
                    maxColumnIndex = dt.Columns.Count;


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
                else if (IsDapperRowOrDictionaryStringObject(value))
                {
                    var collection = value as IEnumerable<IDictionary<string, object>>;

                    var firstTime = true;
                    ICollection<string> keys = null;
                    foreach (var v in collection)
                    {
                        // head
                        if (firstTime)
                        {
                            firstTime = false;
                            if (v == null)
                                continue;

                            keys = v.Keys;
                            maxColumnIndex = keys.Count;
                            if (printHeader)
                            {
                                sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                                var xIndex = xy.Item1;
                                foreach (var key in keys)
                                {
                                    var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                                    sb.Append($"<x:c r=\"{columname}\" t=\"str\">");
                                    sb.Append($"<x:v>{key}");
                                    sb.Append($"</x:v>");
                                    sb.Append($"</x:c>");
                                    xIndex++;
                                }
                                sb.AppendLine($"</x:row>");
                                yIndex++;
                            }
                        }

                        //body
                        {
                            sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                            var xIndex = xy.Item1;
                            foreach (var p in keys)
                            {
                                var cellValue = v[p];
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
                    maxRowIndex = yIndex - 1;
                }
                else if (value is ICollection)
                {
                    var collection = value as ICollection;

                    var props = Helpers.GetSubtypeProperties(collection);

                    maxColumnIndex = props.Length;

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
                    maxRowIndex = yIndex - 1;
                }
                else
                {
                    throw new NotImplementedException($"{value?.GetType().Name} type not implemented,please issue for me : https://github.com/shps951023/MiniExcel/issues");
                }

                // dimension
                {
                    if (maxRowIndex == 0 && maxColumnIndex == 0)
                        dimensionRef = "A1";
                    else if (maxColumnIndex == 1)
                        dimensionRef = $"A{maxRowIndex}";
                    else
                        dimensionRef = $"A1:{Helpers.GetAlphabetColumnName(maxColumnIndex - 1)}{maxRowIndex}";
                }

                defaultFiles[@"xl/worksheets/sheet1.xml"].Xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<dimension ref=""{dimensionRef}""/>
<x:sheetData>{sb.ToString()}</x:sheetData>
</x:worksheet>";
            }

            return defaultFiles;
        }

        private static bool IsDapperRowOrDictionaryStringObject(object value)
        {
            return value is IEnumerable<IDictionary<string, object>>;
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
                        object itemValue = (object)item[p.Name];
                        if (itemValue == null)
                        {
                            yield return v;
                        }
                        else if (p.PropertyType == typeof(Guid) || p.PropertyType == typeof(Guid?))
                            newV = Guid.Parse(itemValue.ToString());
                        else if (p.PropertyType == typeof(DateTime) || p.PropertyType == typeof(DateTime?))
                        {
                            var vs = itemValue.ToString();
                            if (DateTime.TryParse(vs, out var _v))
                                newV = _v;
                            else if (DateTime.TryParseExact(vs, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v2))
                                newV = _v2;
                            else if (double.TryParse(vs, out var _d))
                                newV = DateTimeHelper.FromOADate(_d);
                            else
                                throw new InvalidCastException($"{vs} can't cast to datetime");
                        }
                        else if (p.PropertyType == typeof(bool) || p.PropertyType == typeof(bool?))
                        {
                            var vs = itemValue.ToString();
                            if (vs == "1")
                                newV = true;
                            else if (vs == "0")
                                newV = false;
                            else
                                newV = bool.Parse(vs);
                        }
                        else
                            newV = Convert.ChangeType(itemValue, p.PropertyType);
                        p.SetValue(v, newV);
                    }
                }
                yield return v;
            }
        }

        private static void SaveAsImpl(string path, Dictionary<string, ZipPackageInfo> zipPackageInfos)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false, Utf8WithBom))
                CreteXlsxImpl(zipPackageInfos, archive);
        }
        private static void SaveAsImpl(Stream stream, Dictionary<string, ZipPackageInfo> zipPackageInfos)
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
