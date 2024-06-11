using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MiniExcelLibs.OpenXml.Constants
{
    internal static class ExcelXml
    {
        static ExcelXml()
        {
            DefaultRels = MinifyXml(DefaultRels);
            DefaultWorkbookXml = MinifyXml(DefaultWorkbookXml);
            DefaultStylesXml = MinifyXml(DefaultStylesXml);
            DefaultWorkbookXmlRels = MinifyXml(DefaultWorkbookXmlRels);
            DefaultSheetRelXml = MinifyXml(DefaultSheetRelXml);
            DefaultDrawing = MinifyXml(DefaultDrawing);
        }

        private static string MinifyXml(string xml) => xml.Replace("\r", "").Replace("\n", "").Replace("\t", "");

        internal static readonly string EmptySheetXml = $@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><x:dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>";

        internal static readonly string DefaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

        internal static readonly string DefaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{sheets}}
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""/xl/sharedStrings.xml"" Id=""R3db9602ace778fdb"" />
</Relationships>";

        internal static readonly string NoneStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
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

        internal const string NumFmtsToken = "{{numFmts}}";
        internal const string NumFmtsCountToken = "{{numFmtCount}}";
        internal const string cellXfsToken = "{{cellXfs}}";
        internal const string cellXfsCountToken = "{{cellXfsCount}}";
        internal static readonly string DefaultStylesXml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:numFmts count=""{NumFmtsCountToken}"">
        <x:numFmt numFmtId=""0"" formatCode="""" />
        {NumFmtsToken}
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
    <x:cellXfs count=""{cellXfsCountToken}"">
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
        {cellXfsToken}
    </x:cellXfs>
    <x:cellStyles count=""1"">
        <x:cellStyle name=""Normal"" xfId=""0"" builtinId=""0"" />
    </x:cellStyles>
</x:styleSheet>";

        internal static readonly string DefaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        {{sheets}}
    </x:sheets>
</x:workbook>";

        internal static readonly string DefaultSheetRelXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{format}}
</Relationships>";
        internal static readonly string DefaultDrawing = @"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<xdr:wsDr xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"">
    {{format}}
</xdr:wsDr>";
        internal static readonly string DefaultDrawingXmlRels = @"<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    {{format}}
</Relationships>";

        internal static readonly string DefaultSharedString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";

        internal static readonly string StartTypes = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"" Extension=""bin""/><Default ContentType=""application/xml"" Extension=""xml""/><Default ContentType=""image/jpeg"" Extension=""jpg""/><Default ContentType=""image/png"" Extension=""png""/><Default ContentType=""image/gif"" Extension=""gif""/><Default ContentType=""application/vnd.openxmlformats-package.relationships+xml"" Extension=""rels""/>";
        internal static string ContentType(string contentType, string partName) => $"<Override ContentType=\"{contentType}\" PartName=\"/{partName}\" />";
        internal static readonly string EndTypes = "</Types>";

        internal static string WorksheetRelationship(SheetDto sheetDto)
            => $@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/{sheetDto.Path}"" Id=""{sheetDto.ID}"" />";

        internal static string ImageRelationship(FileDto image)
            => $@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""{image.Path2}"" Id=""{image.ID}"" />";

        internal static string DrawingRelationship(int sheetId)
            => $@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"" Target=""../drawings/drawing{sheetId}.xml"" Id=""drawing{sheetId}"" />";

        internal static string DrawingXml(FileDto file, int fileIndex)
            => $@"<xdr:oneCellAnchor>
        <xdr:from>
            <xdr:col>{file.CellIndex - 1/* why -1 : https://user-images.githubusercontent.com/12729184/150460189-f08ed939-44d4-44e1-be6e-9c533ece6be8.png*/}</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>{file.RowIndex - 1}</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:ext cx=""609600"" cy=""190500"" />
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id=""{fileIndex + 1}"" descr="""" name=""2a3f9147-58ea-4a79-87da-7d6114c4877b"" />
                <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect=""1"" />
                </xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip r:embed=""{file.ID}"" cstate=""print"" />
                <a:stretch>
                    <a:fillRect />
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x=""0"" y=""0"" />
                    <a:ext cx=""0"" cy=""0"" />
                </a:xfrm>
                <a:prstGeom prst=""rect"">
                    <a:avLst />
                </a:prstGeom>
            </xdr:spPr>
        </xdr:pic>
        <xdr:clientData />
    </xdr:oneCellAnchor>";

        internal static string Sheet(SheetDto sheetDto, int sheetId)
            => $@"<x:sheet name=""{ExcelOpenXmlUtils.EncodeXML(sheetDto.Name)}"" sheetId=""{sheetId}""{(string.IsNullOrWhiteSpace(sheetDto.State) ? string.Empty : $" state=\"{sheetDto.State}\"")} r:id=""{sheetDto.ID}"" />";

        internal static string SetupStyleXml(string styleXml, ICollection<ExcelColumnAttribute> columns)
        {
            var sb = new StringBuilder(styleXml);
            var columnsToApply = columns
                .Where(x => !string.IsNullOrWhiteSpace(x.Format))
                .Select(x => new { exf = new ExcelNumberFormat(x.Format), f = x })
                .Where(x => x.exf.IsValid)
                .GroupBy(x => x.f.FormatId)
                .Select(x => x.First()).ToArray();


            var numFmtXmls = columnsToApply.Select((x,i) => $@"<x:numFmt numFmtId=""{166 + i}"" formatCode=""{x.exf.FormatString}"" />").ToArray();
            sb.Replace(NumFmtsToken, string.Join(string.Empty, numFmtXmls));
            sb.Replace(NumFmtsCountToken, (1 + numFmtXmls.Length).ToString());

            var cellXfsXmls = columnsToApply.Select((x, i) =>
            {
                return 
$@"<x:xf numFmtId=""{166 + i}"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
    <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
    <x:protection locked=""1"" hidden=""0"" />
</x:xf>";
            }
            ).ToArray();
            sb.Replace(cellXfsToken, string.Join(string.Empty, cellXfsXmls));
            sb.Replace(cellXfsCountToken, (5 + cellXfsXmls.Length).ToString());
            return sb.ToString();
        }
    }
}
