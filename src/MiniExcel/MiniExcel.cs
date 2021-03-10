namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Xml.Linq;

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

        public static void Create(string filePath, object value, string startCell = "A1", bool printHeader = true)
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
                    var props = type.GetProperties();
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

            CreateXlsxFile(filePath, defaultFiles);
        }

        public static Dictionary<int, Dictionary<int, object>> Query(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                return stream.Query();
            }
        }

        public static Dictionary<int, Dictionary<int, object>> Query(this Stream stream)
        {
            using (var reader = new ExcelOpenXmlReader(stream))
            {
                var d = new Dictionary<int, Dictionary<int, object>>();
                while (reader.Read())
                {
                    var dic = new Dictionary<int, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        var v = reader.GetValue(i);
                        dic.Add(i, v);
                    }
                    d.Add(reader.CurrentRowIndex, dic);
                }
                return d;
            }
        }

        internal static Worksheet GetFirstSheet(Stream stream)
        {
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, UTF8Encoding.UTF8))
            {
                var rows = new Dictionary<int, Dictionary<int, object>>();

                //sharedStrings must in memory cache
                Dictionary<int, string> GetSharedStrings()
                {
                    var sharedStringsEntry = archive.Entries.SingleOrDefault(w => w.FullName == "xl/sharedStrings.xml");
                    var xml = ConvertToString(sharedStringsEntry);
                    var xl = XElement.Parse(xml);
                    var ts = xl.Descendants(ExcelOpenXmlXName.T).Select((s, i) => new { i, v = s.Value?.ToString() })
                         .ToDictionary(s => s.i, s => s.v)
                    ;
                    return ts;
                }

                var sharedStrings = GetSharedStrings();
                
                var rowIndexMaximum = int.MinValue;
                var columnIndexMaximum = int.MinValue;

                //notice: for performance just read first one and no care the order
                var firstSheetEntry = archive.Entries.First(w => w.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase));
                {
                    var xml = ConvertToString(firstSheetEntry);
                    var xl = XElement.Parse(xml);

                    foreach (var row in xl.Descendants(ExcelOpenXmlXName.Row))
                    {
                        //
                        var datarow = new Dictionary<int, object>();
                        {
                            var r = row.Attribute("r")?.Value?.ToString();

                            var rowIndex = int.MinValue;
                            if (int.TryParse(r, out var _rowIndex))
                                rowIndex = _rowIndex - 1; // The row attribute is 1 - based				
                            rowIndexMaximum = Math.Max(rowIndexMaximum, rowIndex);

                            rows.Add(rowIndex, datarow);
                        }

                        foreach (var cell in row.Descendants(ExcelOpenXmlXName.C))
                        {
                            var t = cell.Attribute("t")?.Value?.ToString();
                            var v = cell.Descendants(ExcelOpenXmlXName.V).SingleOrDefault()?.Value;
                            if (t == "s")
                            {
                                if (!string.IsNullOrEmpty(v))
                                    v = sharedStrings[int.Parse(v)];
                            }

                            var r = cell.Attribute("r")?.Value?.ToString();
                            {
                                var cellIndex = ExcelOpenXmlUtils.GetCellColumnIndex(r) - 1;
                                columnIndexMaximum = Math.Max(columnIndexMaximum, cellIndex);

                                datarow.Add(cellIndex, v);
                            }
                        }
                        
                    }
                }

                return new Worksheet {
                    Rows = rows,
                    FieldCount = columnIndexMaximum + 1,
                    RowCount = rowIndexMaximum + 1
                };
            }
        }

        private static string ConvertToString(ZipArchiveEntry entry)
        {
            using (var eStream = entry.Open())
            using (var reader = new StreamReader(eStream))
                return reader.ReadToEnd();
        }

        private readonly static UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);
        private static void CreateXlsxFile(string path, Dictionary<string, ZipPackageInfo> zipPackageInfos)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false, UTF8Encoding.UTF8))
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
}
