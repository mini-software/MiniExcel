using MiniExcelLibs.Attributes;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MiniExcelLibs.OpenXml.Styles {

    public class DefaultSheetStyleBuilder : ISheetStyleBuilder
    {
        private const int startUpNumFmts = 1;
        private const string NumFmtsToken = "{{numFmts}}";
        private const string NumFmtsCountToken = "{{numFmtCount}}";

        private const int startUpCellXfs = 5;
        private const string cellXfsToken = "{{cellXfs}}";
        private const string cellXfsCountToken = "{{cellXfsCount}}";

        internal static readonly string DefaultStylesXml = ExcelOpenXmlUtils.MinifyXml
        ( $@"
            <?xml version=""1.0"" encoding=""utf-8""?>
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
            </x:styleSheet>"
        );

        public string Build( ICollection<ExcelColumnAttribute> columns )
        {
            const int numFmtIndex = 166;

            var sb = new StringBuilder( DefaultStylesXml );
            var columnsToApply = SheetStyleBuilderHelper.GenerateStyleIds( startUpCellXfs, columns );

            var numFmts = columnsToApply.Select( ( x, i ) =>
            {
                return new
                {
                    numFmt = $@"<x:numFmt numFmtId=""{numFmtIndex + i}"" formatCode=""{x.Format}"" />",

                    cellXfs = $@"<x:xf numFmtId=""{numFmtIndex + i}"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
    <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
    <x:protection locked=""1"" hidden=""0"" />
</x:xf>"
                };
            } ).ToArray();

            sb.Replace( NumFmtsToken, string.Join( string.Empty, numFmts.Select( x => x.numFmt ) ) );
            sb.Replace( NumFmtsCountToken, (startUpNumFmts + numFmts.Length).ToString() );

            sb.Replace( cellXfsToken, string.Join( string.Empty, numFmts.Select( x => x.cellXfs ) ) );
            sb.Replace( cellXfsCountToken, (5 + numFmts.Length).ToString() );
            return sb.ToString();
        }
    }

}
