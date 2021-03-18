using MiniExcelLibs.Zip;
using System.Collections.Generic;

namespace MiniExcelLibs.OpenXml
{
    internal static class DefualtXml
    {
        internal const string DefaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""/xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

        internal const string DefaultSheetXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheetData>
    </x:sheetData>
</x:worksheet>";
        internal const string DefaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/xl/worksheets/sheet1.xml"" Id=""R1274d0d920f34a32"" />
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
</Relationships>";

        internal const string DefaultStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
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

        internal const string DefaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        <x:sheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" name=""Sheet1"" sheetId=""1"" r:id=""R1274d0d920f34a32"" />
    </x:sheets>
</x:workbook>";
        public static Dictionary<string, ZipPackageInfo> GetDefaultFiles() => new Dictionary<string, ZipPackageInfo>()
        {
            { @"_rels/.rels",new ZipPackageInfo(DefualtXml.DefaultRels, "application/vnd.openxmlformats-package.relationships+xml")},
            { @"xl/_rels/workbook.xml.rels",new ZipPackageInfo(DefualtXml.DefaultWorkbookXmlRels, "application/vnd.openxmlformats-package.relationships+xml")},
            { @"xl/styles.xml",new ZipPackageInfo(DefualtXml.DefaultStylesXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")},
            { @"xl/workbook.xml",new ZipPackageInfo(DefualtXml.DefaultWorkbookXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")},
            { @"xl/worksheets/sheet1.xml",new ZipPackageInfo(DefualtXml.DefaultSheetXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")},
        };
    }
}
