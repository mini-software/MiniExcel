using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly Dictionary<string, ZipPackageInfo> _zipDictionary = new Dictionary<string, ZipPackageInfo>();

        static ExcelOpenXmlSheetWriter()
        {
            _defaultRels = MinifyXml(_defaultRels);
            _defaultWorkbookXml = MinifyXml(_defaultWorkbookXml);
            _defaultStylesXml = MinifyXml(_defaultStylesXml);
            _defaultWorkbookXmlRels = MinifyXml(_defaultWorkbookXmlRels);
            _defaultSheetRelXml = MinifyXml(_defaultSheetRelXml);
            _defaultDrawing = MinifyXml(_defaultDrawing);
        }

        private static readonly string _defaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

        private static readonly string _defaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{sheets}}
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
</Relationships>";

        private static readonly string _noneStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
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
	   <x:xf />
	   <x:xf />
        <x:xf numFmtId=""14"" applyNumberFormat=""1"" />
    </x:cellXfs>
</x:styleSheet>";

        private static readonly string _defaultStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:numFmts count=""1"">
        <x:numFmt numFmtId=""0"" formatCode="""" />
    </x:numFmts>
    <x:fonts count=""2"">
        <x:font>
            <x:vertAlign val=""baseline"" />
            <x:sz val=""11"" />
            <x:color rgb=""FF000000"" />
            <x:name val=""Calibri"" />
            <x:family val=""2"" />
        </x:font>
        <x:font>
            <x:vertAlign val=""baseline"" />
            <x:sz val=""11"" />
            <x:color rgb=""FFFFFFFF"" />
            <x:name val=""Calibri"" />
            <x:family val=""2"" />
        </x:font>
    </x:fonts>
    <x:fills count=""3"">
        <x:fill>
            <x:patternFill patternType=""none"" />
        </x:fill>
        <x:fill>
            <x:patternFill patternType=""gray125"" />
        </x:fill>
        <x:fill>
            <x:patternFill patternType=""solid"">
                <x:fgColor rgb=""284472C4"" />
            </x:patternFill>
        </x:fill>
    </x:fills>
    <x:borders count=""2"">
        <x:border diagonalUp=""0"" diagonalDown=""0"">
            <x:left style=""none"">
                <x:color rgb=""FF000000"" />
            </x:left>
            <x:right style=""none"">
                <x:color rgb=""FF000000"" />
            </x:right>
            <x:top style=""none"">
                <x:color rgb=""FF000000"" />
            </x:top>
            <x:bottom style=""none"">
                <x:color rgb=""FF000000"" />
            </x:bottom>
            <x:diagonal style=""none"">
                <x:color rgb=""FF000000"" />
            </x:diagonal>
        </x:border>
        <x:border diagonalUp=""0"" diagonalDown=""0"">
            <x:left style=""thin"">
                <x:color rgb=""FF000000"" />
            </x:left>
            <x:right style=""thin"">
                <x:color rgb=""FF000000"" />
            </x:right>
            <x:top style=""thin"">
                <x:color rgb=""FF000000"" />
            </x:top>
            <x:bottom style=""thin"">
                <x:color rgb=""FF000000"" />
            </x:bottom>
            <x:diagonal style=""none"">
                <x:color rgb=""FF000000"" />
            </x:diagonal>
        </x:border>
    </x:borders>
    <x:cellStyleXfs count=""3"">
        <x:xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""0"" applyAlignment=""1"" applyProtection=""1"">
            <x:protection locked=""1"" hidden=""0"" />
        </x:xf>
        <x:xf numFmtId=""14"" fontId=""1"" fillId=""2"" borderId=""1"" applyNumberFormat=""1"" applyFill=""0"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
            <x:protection locked=""1"" hidden=""0"" />
        </x:xf>
        <x:xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""1"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
            <x:protection locked=""1"" hidden=""0"" />
        </x:xf>
    </x:cellStyleXfs>
    <x:cellXfs count=""4"">
        <x:xf></x:xf>
        <x:xf numFmtId=""0"" fontId=""1"" fillId=""2"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""0"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
            <x:alignment horizontal=""left"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
            <x:protection locked=""1"" hidden=""0"" />
        </x:xf>
        <x:xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
            <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
            <x:protection locked=""1"" hidden=""0"" />
        </x:xf>
        <x:xf numFmtId=""14"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
            <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
            <x:protection locked=""1"" hidden=""0"" />
        </x:xf>
        <x:xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyBorder=""1"" applyAlignment=""1"">
            <x:alignment horizontal=""fill""/>
        </x:xf>
    </x:cellXfs>
    <x:cellStyles count=""1"">
        <x:cellStyle name=""Normal"" xfId=""0"" builtinId=""0"" />
    </x:cellStyles>
</x:styleSheet>";

        private static readonly string _defaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        {{sheets}}
    </x:sheets>
</x:workbook>";

        private static readonly string _defaultSheetRelXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{format}}
</Relationships>";
        private static readonly string _defaultDrawing = @"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<xdr:wsDr xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
    {{format}}
</xdr:wsDr>";
        private static readonly string _defaultDrawingXmlRels = @"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{format}}
</Relationships>";

        private static readonly string _defaultSharedString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";
        private static string MinifyXml(string xml) => xml.Replace("\r", "").Replace("\n", "").Replace("\t", "");

        internal void GenerateDefaultOpenXml()
        {
            CreateZipEntry("_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", ExcelOpenXmlSheetWriter._defaultRels);
            CreateZipEntry("xl/sharedStrings.xml", "application/vnd.openxmlformats-package.relationships+xml", ExcelOpenXmlSheetWriter._defaultSharedString);
        }

        private void CreateZipEntry(string path,string contentType,string content)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path);
            using (var zipStream = entry.Open())
            using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom,_configuration.BufferSize))
                writer.Write(content);
            if(!string.IsNullOrEmpty(contentType))
                _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
        }

        private void CreateZipEntry(string path, byte[] content)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path);
            using (var zipStream = entry.Open())
                zipStream.Write(content,0, content.Length);
        }
    }
}
