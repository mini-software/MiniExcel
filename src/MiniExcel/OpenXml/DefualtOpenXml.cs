using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace MiniExcelLibs.OpenXml
{
    internal static class DefualtOpenXml
    {
        private readonly static UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);

        internal static string DefaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""/xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

        internal static string DefaultSheetXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheetData>
    </x:sheetData>
</x:worksheet>";
        internal static string DefaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/xl/worksheets/sheet1.xml"" Id=""R1274d0d920f34a32"" />
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
</Relationships>";

        internal static string DefaultStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:fonts>
        <x:font />
    </x:fonts>
    <x:fills>
        <x:fill />
    </x:fills>
    <x:borders>
        <x:border />
    </x:borders>
    <x:cellStyleXfs>
        <x:xf />
    </x:cellStyleXfs>
    <x:cellXfs>
        <x:xf />
        <x:xf numFmtId=""14"" applyNumberFormat=""1"" />
    </x:cellXfs>
</x:styleSheet>";

        internal static string DefaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        <x:sheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" name=""Sheet1"" sheetId=""1"" r:id=""R1274d0d920f34a32"" />
    </x:sheets>
</x:workbook>";

        static DefualtOpenXml()
        {
            DefaultRels = MinifyXml(DefaultRels);
            DefaultWorkbookXml = MinifyXml(DefaultWorkbookXml);
            DefaultStylesXml = MinifyXml(DefaultStylesXml);
            DefaultWorkbookXmlRels = MinifyXml(DefaultWorkbookXmlRels);
            DefaultSheetXml = MinifyXml(DefaultSheetXml);
        }

        private static string MinifyXml(string xml) => xml
            //.Replace("    ", "").Replace("\r", "").Replace("\n", "").Replace("\t", "")
            ;

        //TODO:read from static generated file looks like more better?
        internal static Dictionary<string,ZipPackageInfo> GenerateDefaultOpenXml(ZipArchive archive)
        {
            var defaults = new Dictionary<string, Tuple<string,string>>()
            {
                { @"_rels/.rels", new Tuple<string,string>(DefualtOpenXml.DefaultRels, "application/vnd.openxmlformats-package.relationships+xml")},
                { @"xl/_rels/workbook.xml.rels", new Tuple<string,string>(DefualtOpenXml.DefaultWorkbookXmlRels, "application/vnd.openxmlformats-package.relationships+xml")},
                { @"xl/styles.xml", new Tuple<string,string>(DefualtOpenXml.DefaultStylesXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")},
                { @"xl/workbook.xml", new Tuple<string,string>(DefualtOpenXml.DefaultWorkbookXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")},
                //{ @"xl/worksheets/sheet1.xml",new Tuple<string,string>(DefualtOpenXml.DefaultSheetXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")},
            };

            var zps = new Dictionary<string, ZipPackageInfo>();
            foreach (var p in defaults)
            {
                ZipArchiveEntry entry = archive.CreateEntry(p.Key);
                using (var zipStream = entry.Open())
                using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    writer.Write(p.Value.Item1.ToString());

                zps.Add(p.Key, new ZipPackageInfo(entry, p.Value.Item2));
            }
            return zps;
        }
    }
}
