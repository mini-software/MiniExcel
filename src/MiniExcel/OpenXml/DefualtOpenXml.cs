using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace MiniExcelLibs.OpenXml
{
    internal static class DefualtOpenXml
    {
        private readonly static UTF8Encoding Utf8WithBom = new UTF8Encoding(true);

        private static string DefaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""/xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

        private static string DefaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{sheets}}
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
</Relationships>";

        private static string NoneStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
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

        private static string DefaultStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
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
    <x:cellXfs count=""3"">
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
    </x:cellXfs>
    <x:cellStyles count=""1"">
        <x:cellStyle name=""Normal"" xfId=""0"" builtinId=""0"" />
    </x:cellStyles>
</x:styleSheet>";

        private static string DefaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        {{sheets}}
    </x:sheets>
</x:workbook>";

        private static Dictionary<string, XmlDocument> Xmls = new Dictionary<string, XmlDocument>();

        private static readonly XmlNamespaceManager _ns;
        static DefualtOpenXml()
        {
            DefaultRels = MinifyXml(DefaultRels);
            DefaultWorkbookXml = MinifyXml(DefaultWorkbookXml);
            DefaultStylesXml = MinifyXml(DefaultStylesXml);
            DefaultWorkbookXmlRels = MinifyXml(DefaultWorkbookXmlRels);

            _ns = new XmlNamespaceManager(new NameTable());
            _ns.AddNamespace("x", Config.SpreadsheetmlXmlns);

            Xmls.Add("DefaultStylesXml", GetXmlDocument(DefaultStylesXml));
        }

        private static XmlDocument GetXmlDocument(string xml)
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);
            return doc;
        }

        private static string MinifyXml(string xml) => xml
        //.Replace("    ", "").Replace("\r", "").Replace("\n", "").Replace("\t", "")
        ;

        //TODO:read from static generated file looks like more better?
        internal static Dictionary<string, ZipPackageInfo> GenerateDefaultOpenXml(ZipArchive archive, IEnumerable<string> sheetNames, OpenXmlConfiguration configuration)
        {
            var defaults = new Dictionary<string, Tuple<string, string>>()
            {
                { @"_rels/.rels", new Tuple<string,string>(DefualtOpenXml.DefaultRels, "application/vnd.openxmlformats-package.relationships+xml")},
            };

            // styles.xml
            {
                var styleXml = string.Empty;

                if (configuration.TableStyles == TableStyles.None)
                {
                    styleXml = NoneStylesXml;
                }
                else if (configuration.TableStyles == TableStyles.Default)
                {
                    styleXml = DefaultStylesXml;
                }

                defaults.Add(@"xl/styles.xml", new Tuple<string, string>(styleXml, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
            }

            // workbook.xml 、 workbookRelsXml
            {
                var workbookXml = new StringBuilder();
                var workbookRelsXml = new StringBuilder();

                var sheetId = 0;
                foreach (var sheetName in sheetNames)
                {
                    sheetId++;
                    var id = $"R{Guid.NewGuid().ToString("N")}";
                    workbookXml.AppendLine($@"<x:sheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" name=""{sheetName}"" sheetId=""{sheetId}"" r:id=""{id}"" />");
                    workbookRelsXml.AppendLine($@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/xl/worksheets/sheet{sheetId}.xml"" Id=""{id}"" />");
                }
                defaults.Add(@"xl/workbook.xml", new Tuple<string, string>(
                    DefualtOpenXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString())
                    , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
                );
                defaults.Add(@"xl/_rels/workbook.xml.rels", new Tuple<string, string>(
                    DefualtOpenXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString())
                    , "application/vnd.openxmlformats-package.relationships+xml")
                );
            }

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
