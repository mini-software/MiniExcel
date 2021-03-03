namespace MiniExcel
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Xml;
    using System.Xml.Linq;
    public static class MiniExcelHelper
    {
        internal static Dictionary<string, ZipPackageInfo> DefaultFilesTree => new Dictionary<string, ZipPackageInfo>()
        {
            { @"_rels/.rels",new ZipPackageInfo(DefualtXml.defaultRels, "application/vnd.openxmlformats-package.relationships+xml")},
            { @"xl/_rels/workbook.xml.rels",new ZipPackageInfo(DefualtXml.defaultWorkbookXmlRels, "application/vnd.openxmlformats-package.relationships+xml")},
            { @"xl/styles.xml",new ZipPackageInfo(DefualtXml.defaultStylesXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")},
            { @"xl/workbook.xml",new ZipPackageInfo(DefualtXml.defaultWorkbookXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")},
            { @"xl/worksheets/sheet1.xml",new ZipPackageInfo(DefualtXml.defaultSheetXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")},
        };

        private static FileStream CreateZipFileStream(string path, Dictionary<string, object> filesTree)
        {
            var utf8WithBom = new System.Text.UTF8Encoding(true);  // 用true来指定包含bom
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false, UTF8Encoding.UTF8))
                {
                    foreach (var fileTree in filesTree)
                    {
                        ZipArchiveEntry entry = archive.CreateEntry(fileTree.Key);
                        using (var zipStream = entry.Open())
                        {
                            //var bytes = utf8WithBom.GetBytes(fileTree.Value.ToString());
                            //zipStream.Write(bytes, 0, bytes.Length);

                            using (StreamWriter writer = new StreamWriter(zipStream, utf8WithBom))
                            {
                                writer.Write(fileTree.Value.ToString()); //entry contents "baz123"
                            }
                        }

                    }
                }
                return stream;
            }
        }

        public static void Create(string path, object value, string startCell = "A1", bool printHeader = true)
        {
            var xy = XlsxUtils.ConvertCellToXY(startCell);

            var filesTree = DefaultFilesTree;
            {
                var sb = new StringBuilder();

                var yIndex = xy.Item2;

                if (value is System.Collections.ICollection)
                {
                    var _vs = value as System.Collections.ICollection;
                    object firstValue = null;
                    {
                        foreach (var v in _vs)
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
                            var columname = XlsxUtils.ConvertXyToCell(xIndex, yIndex);
                            sb.Append($"<x:c r=\"{columname}\" t=\"str\">");
                            sb.Append($"<x:v>{p.Name}");
                            sb.Append($"</x:v>");
                            sb.Append($"</x:c>");
                            xIndex++;
                        }
                        sb.AppendLine($"</x:row>");
                        yIndex++;
                    }

                    foreach (var v in _vs)
                    {
                        sb.AppendLine($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (var p in props)
                        {
                            var cellValue = p.GetValue(v);
                            var cellValueStr = XlsxUtils.GetValue(cellValue);
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
                            var columname = XlsxUtils.ConvertXyToCell(xIndex, yIndex);
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

                filesTree[@"xl/worksheets/sheet1.xml"].Xml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<x:sheetData>{sb.ToString()}</x:sheetData>
</x:worksheet>";
            }

            CreateXlsxFile(path, filesTree);
        }

        //public static Dictionary<string, object> Read(string fileName)
        //{
        //    var parsedCells = new Dictionary<string, object>();
        //    using (Package xlsxPackage = Package.Open(fileName, FileMode.Open, FileAccess.Read))
        //    {
        //        var allParts = xlsxPackage.GetParts();


        //        var worksheetElement = GetFirstWorksheet(allParts);
        //        var cells = from c in worksheetElement.Descendants(ExcelNamespaces.excelNamespace + "c")
        //                    select c;

        //        var sharedStrings = GetSharedStrings(allParts);
        //        foreach (XElement cell in cells)
        //        {
        //            var r = cell.Attribute("r");
        //            {
        //                var cellPosition = r.Value;
        //                var v = cell.Descendants(ExcelNamespaces.excelNamespace + "v").SingleOrDefault()?.Value;
        //                var t = cell.Attribute("t")?.Value;
        //                if (t == "s")
        //                {
        //                    parsedCells.Add(cellPosition, sharedStrings[Convert.ToInt32(v)]);
        //                }
        //                else
        //                {
        //                    parsedCells.Add(cellPosition, v);
        //                }
        //            }

        //        }
        //    }

        //    return parsedCells;
        //}

        //private static Dictionary<int, string> GetSharedStrings(PackagePartCollection allParts)
        //{
        //    var sharedStringsPart = (from part in allParts
        //                             where part.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        //                             select part).SingleOrDefault();
        //    if (sharedStringsPart == null)
        //        return null;


        //    Dictionary<int, string> sharedStrings = new Dictionary<int, string>();
        //    var sharedStringsElement = XElement.Load(XmlReader.Create(sharedStringsPart.GetStream()));
        //    IEnumerable<XElement> sharedStringsElements = from s in sharedStringsElement.Descendants(ExcelNamespaces.excelNamespace + "t")
        //                                                  select s;
        //    int Counter = 0;
        //    foreach (XElement sharedString in sharedStringsElements)
        //    {
        //        sharedStrings.Add(Counter, sharedString.Value);
        //        Counter++;
        //    }
        //    return sharedStrings;
        //}

        //private static XElement GetFirstWorksheet(PackagePartCollection allParts)
        //{
        //    PackagePart worksheetPart = (from part in allParts
        //                                 where part.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        //                                 select part).FirstOrDefault();

        //    return XElement.Load(XmlReader.Create(worksheetPart.GetStream()));
        //}

        private readonly static UTF8Encoding _utf8WithBom = new System.Text.UTF8Encoding(true);
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
                    using (StreamWriter writer = new StreamWriter(zipStream, _utf8WithBom))
                        writer.Write(sb.ToString());
                }

                foreach (var p in zipPackageInfos)
                {
                    ZipArchiveEntry entry = archive.CreateEntry(p.Key);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, _utf8WithBom))
                        writer.Write(p.Value.Xml.ToString());
                }
            }
        }
    }

    internal static class XlsxUtils
    {
        internal static string GetValue(object value) => value == null ? "" : value.ToString().Replace("<", "&lt;").Replace(">", "&gt;");

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        internal static string ConvertXyToCell(Tuple<int, int> xy)
        {
            return ConvertXyToCell(xy.Item1, xy.Item2);
        }

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        internal static string ConvertXyToCell(int x, int y)
        {
            int dividend = x;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return $"{columnName}{y}";
        }

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        internal static Tuple<int, int> ConvertCellToXY(string cell)
        {
            const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const int mode = 26;

            var x = 0;
            var cellLetter = GetCellLetter(cell);
            //AA=27,ZZ=702
            for (int i = 0; i < cellLetter.Length; i++)
                x = x * mode + keys.IndexOf(cellLetter[i]);

            var cellNumber = GetCellNumber(cell);
            return Tuple.Create(x, int.Parse(cellNumber));
        }

        internal static string GetCellNumber(string cell)
        {
            string cellNumber = string.Empty;
            for (int i = 0; i < cell.Length; i++)
            {
                if (Char.IsDigit(cell[i]))
                    cellNumber += cell[i];
            }
            return cellNumber;
        }

        internal static string GetCellLetter(string cell)
        {
            string GetCellLetter = string.Empty;
            for (int i = 0; i < cell.Length; i++)
            {
                if (Char.IsLetter(cell[i]))
                    GetCellLetter += cell[i];
            }
            return GetCellLetter;
        }
    }
}
