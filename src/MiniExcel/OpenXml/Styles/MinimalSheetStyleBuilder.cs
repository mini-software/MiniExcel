using MiniExcelLibs.Attributes;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MiniExcelLibs.OpenXml.Styles {
    public class MinimalSheetStyleBuilder : ISheetStyleBuilder
    {
        private const int startUpNumFmts = 1;
        private const string NumFmtsToken = "{{numFmts}}";
        private const string NumFmtsCountToken = "{{numFmtCount}}";

        private const int startUpCellXfs = 5;
        private const string cellXfsToken = "{{cellXfs}}";
        private const string cellXfsCountToken = "{{cellXfsCount}}";

        internal static readonly string NoneStylesXml = ExcelOpenXmlUtils.MinifyXml
        ( $@"
            <?xml version=""1.0"" encoding=""utf-8""?>
            <x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
                <x:numFmts count=""{NumFmtsCountToken}"">
                    <x:numFmt numFmtId=""0"" formatCode="""" />
                    {NumFmtsToken}
                </x:numFmts>
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
                <x:cellXfs count=""{cellXfsCountToken}"">
                    <x:xf />
                    <x:xf />
                    <x:xf />
                    <x:xf numFmtId=""14"" applyNumberFormat=""1"" />
                    <x:xf />
                    {cellXfsToken}
                </x:cellXfs>
            </x:styleSheet>"
        );
        
        public string Build( ICollection<ExcelColumnAttribute> columns )
        {
            const int numFmtIndex = 166;

            var sb = new StringBuilder( NoneStylesXml );
            var columnsToApply = SheetStyleBuilderHelper.GenerateStyleIds( startUpCellXfs, columns );

            var numFmts = columnsToApply.Select( ( x, i ) => {
                return new {
                    numFmt = $@"<x:numFmt numFmtId=""{numFmtIndex + i}"" formatCode=""{x.Format}"" />",
                    cellXfs = $@"<x:xf numFmtId=""{numFmtIndex + i}"" applyNumberFormat=""1"" />"
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
