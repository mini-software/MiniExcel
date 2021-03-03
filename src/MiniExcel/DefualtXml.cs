namespace MiniExcel
{
    internal static class DefualtXml
    {
        internal const string defaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""/xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

        internal const string defaultSheetXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheetData>
    </x:sheetData>
</x:worksheet>";
        internal const string defaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/xl/worksheets/sheet1.xml"" Id=""R1274d0d920f34a32"" />
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
</Relationships>";

        internal const string defaultStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
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

        internal const string defaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        <x:sheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" name=""Sheet1"" sheetId=""1"" r:id=""R1274d0d920f34a32"" />
    </x:sheets>
</x:workbook>";

        internal const string defaultContent_TypesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"" />
    <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
    <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"" />
    <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" />
</Types>";
    }
}
