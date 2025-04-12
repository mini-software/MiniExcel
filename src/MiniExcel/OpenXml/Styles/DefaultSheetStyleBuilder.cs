using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml.Styles
{
    internal class DefaultSheetStyleBuilder : SheetStyleBuilderBase
    {
        internal static readonly SheetStyleElementInfos GenerateElementInfos = new SheetStyleElementInfos
        {
            NumFmtCount = 0,//默认的NumFmt数量是0，但是会有根据ColumnsToApply动态生成的NumFmt
            FontCount = 2,
            FillCount = 3,
            BorderCount = 2,
            CellStyleXfCount = 3,
            CellXfCount = 5
        };

        private readonly SheetStyleBuildContext _context;

        public DefaultSheetStyleBuilder(SheetStyleBuildContext context) : base(context)
        {
            _context = context;
        }

        protected override SheetStyleElementInfos GetGenerateElementInfos()
        {
            return GenerateElementInfos;
        }

        protected override void GenerateNumFmt()
        {
            const int numFmtIndex = 166;

            var index = 0;
            foreach (var item in _context.ColumnsToApply)
            {
                index++;

                /*
                 * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{x.Format}"
                 */
                _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI);
                _context.NewXmlWriter.WriteAttributeString("numFmtId", (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString());
                _context.NewXmlWriter.WriteAttributeString("formatCode", item.Format);
                _context.NewXmlWriter.WriteFullEndElement();
            }
        }

        protected override async Task GenerateNumFmtAsync()
        {
            const int numFmtIndex = 166;
            var index = 0;
            foreach (var item in _context.ColumnsToApply)
            {
                index++;

                /*
                 * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{x.Format}"
                 */
                await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI);
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString()); ;
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "formatCode", null, item.Format);
                await _context.NewXmlWriter.WriteFullEndElementAsync();
            }
        }

        protected override void GenerateFont()
        {
            /*
             * <x:font>
             *     <x:vertAlign val="baseline" />
             *     <x:sz val="11" />
             *     <x:color rgb="FF000000" />
             *     <x:name val="Calibri" />
             *     <x:family val="2" />
             * </x:font>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "vertAlign", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "baseline");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "sz", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "11");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "name", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "Calibri");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "family", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "2");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:font>
             *     <x:vertAlign val="baseline" />
             *     <x:sz val="11" />
             *     <x:color rgb="FFFFFFFF" />
             *     <x:name val="Calibri" />
             *     <x:family val="2" />
             * </x:font>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "vertAlign", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "baseline");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "sz", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "11");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FFFFFFFF");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "name", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "Calibri");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "family", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("val", "2");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
        }

        protected override async Task GenerateFontAsync()
        {
            /*
            * <x:font>
            *     <x:vertAlign val="baseline" />
            *     <x:sz val="11" />
            *     <x:color rgb="FF000000" />
            *     <x:name val="Calibri" />
            *     <x:family val="2" />
            * </x:font>
            */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "vertAlign", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "baseline");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "sz", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "11");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "name", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "Calibri");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "family", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "2");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:font>
             *     <x:vertAlign val="baseline" />
             *     <x:sz val="11" />
             *     <x:color rgb="FFFFFFFF" />
             *     <x:name val="Calibri" />
             *     <x:family val="2" />
             * </x:font>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "vertAlign", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "baseline");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "sz", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "11");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FFFFFFFF");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "name", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "Calibri");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "family", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "2");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
        }

        protected override void GenerateFill()
        {
            /*
             * <x:fill>
             *     <x:patternFill patternType="none" />
             * </x:fill>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("patternType", "none");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:fill>
             *     <x:patternFill patternType="gray125" />
             * </x:fill>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("patternType", "gray125");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:fill>
             *     <x:patternFill patternType="solid">
             *         <x:fgColor rgb="284472C4" />
             *     </x:patternFill>
             * </x:fill>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("patternType", "solid");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fgColor", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "284472C4");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
        }

        protected override async Task GenerateFillAsync()
        {
            /*
            * <x:fill>
            *     <x:patternFill patternType="none" />
            * </x:fill>
            */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "patternType", null, "none");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:fill>
             *     <x:patternFill patternType="gray125" />
             * </x:fill>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "patternType", null, "gray125");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:fill>
             *     <x:patternFill patternType="solid">
             *         <x:fgColor rgb="284472C4" />
             *     </x:patternFill>
             * </x:fill>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "patternType", null, "solid");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fgColor", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "284472C4");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
        }

        protected override void GenerateBorder()
        {
            /*
             * <x:border diagonalUp="0" diagonalDown="0">
             *     <x:left style="none">
             *         <x:color rgb="FF000000" />
             *     </x:left>
             *     <x:right style="none">
             *         <x:color rgb="FF000000" />
             *     </x:right>
             *     <x:top style="none">
             *         <x:color rgb="FF000000" />
             *     </x:top>
             *     <x:bottom style="none">
             *         <x:color rgb="FF000000" />
             *     </x:bottom>
             *     <x:diagonal style="none">
             *         <x:color rgb="FF000000" />
             *     </x:diagonal>
             * </x:border>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("diagonalUp", "0");
            _context.NewXmlWriter.WriteAttributeString("diagonalDown", "0");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "left", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "none");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "right", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "none");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "top", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "none");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "bottom", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "none");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "diagonal", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "none");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:border diagonalUp="0" diagonalDown="0">
             *     <x:left style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:left>
             *     <x:right style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:right>
             *     <x:top style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:top>
             *     <x:bottom style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:bottom>
             *     <x:diagonal style="none">
             *         <x:color rgb="FF000000" />
             *     </x:diagonal>
             * </x:border>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("diagonalUp", "0");
            _context.NewXmlWriter.WriteAttributeString("diagonalDown", "0");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "left", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "thin");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "right", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "thin");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "top", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "thin");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "bottom", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "thin");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "diagonal", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("style", "none");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("rgb", "FF000000");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
        }

        protected override async Task GenerateBorderAsync()
        {
            /*
             * <x:border diagonalUp="0" diagonalDown="0">
             *     <x:left style="none">
             *         <x:color rgb="FF000000" />
             *     </x:left>
             *     <x:right style="none">
             *         <x:color rgb="FF000000" />
             *     </x:right>
             *     <x:top style="none">
             *         <x:color rgb="FF000000" />
             *     </x:top>
             *     <x:bottom style="none">
             *         <x:color rgb="FF000000" />
             *     </x:bottom>
             *     <x:diagonal style="none">
             *         <x:color rgb="FF000000" />
             *     </x:diagonal>
             * </x:border>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalUp", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalDown", null, "0");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "left", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "right", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "top", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "bottom", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "diagonal", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:border diagonalUp="0" diagonalDown="0">
             *     <x:left style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:left>
             *     <x:right style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:right>
             *     <x:top style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:top>
             *     <x:bottom style="thin">
             *         <x:color rgb="FF000000" />
             *     </x:bottom>
             *     <x:diagonal style="none">
             *         <x:color rgb="FF000000" />
             *     </x:diagonal>
             * </x:border>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalUp", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalDown", null, "0");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "left", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "right", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "top", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "bottom", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "diagonal", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
        }

        protected override void GenerateCellStyleXf()
        {
            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyFill="1" applyBorder="0" applyAlignment="1" applyProtection="1">
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "0");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteAttributeString("applyFill", "1");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "0");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("locked", "1");
            _context.NewXmlWriter.WriteAttributeString("hidden", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:xf numFmtId="14" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "14");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 2}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteAttributeString("applyFill", "0");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("locked", "1");
            _context.NewXmlWriter.WriteAttributeString("hidden", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "0");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteAttributeString("applyFill", "1");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("locked", "1");
            _context.NewXmlWriter.WriteAttributeString("hidden", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();
        }

        protected override async Task GenerateCellStyleXfAsync()
        {
            /*
            * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyFill="1" applyBorder="0" applyAlignment="1" applyProtection="1">
            *     <x:protection locked="1" hidden="0" />
            * </x:xf>
            */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:xf numFmtId="14" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 2}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();
        }

        protected override void GenerateCellXf()
        {
            /*
             * <x:xf></x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:alignment horizontal="left" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "0");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 2}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("xfId", "0");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteAttributeString("applyFill", "0");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("horizontal", "left");
            _context.NewXmlWriter.WriteAttributeString("vertical", "bottom");
            _context.NewXmlWriter.WriteAttributeString("textRotation", "0");
            _context.NewXmlWriter.WriteAttributeString("wrapText", "0");
            _context.NewXmlWriter.WriteAttributeString("indent", "0");
            _context.NewXmlWriter.WriteAttributeString("relativeIndent", "0");
            _context.NewXmlWriter.WriteAttributeString("justifyLastLine", "0");
            _context.NewXmlWriter.WriteAttributeString("shrinkToFit", "0");
            _context.NewXmlWriter.WriteAttributeString("readingOrder", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("locked", "1");
            _context.NewXmlWriter.WriteAttributeString("hidden", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "0");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("xfId", "0");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteAttributeString("applyFill", "1");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("horizontal", "general");
            _context.NewXmlWriter.WriteAttributeString("vertical", "bottom");
            _context.NewXmlWriter.WriteAttributeString("textRotation", "0");
            _context.NewXmlWriter.WriteAttributeString("wrapText", "0");
            _context.NewXmlWriter.WriteAttributeString("indent", "0");
            _context.NewXmlWriter.WriteAttributeString("relativeIndent", "0");
            _context.NewXmlWriter.WriteAttributeString("justifyLastLine", "0");
            _context.NewXmlWriter.WriteAttributeString("shrinkToFit", "0");
            _context.NewXmlWriter.WriteAttributeString("readingOrder", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("locked", "1");
            _context.NewXmlWriter.WriteAttributeString("hidden", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "14");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("xfId", "0");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteAttributeString("applyFill", "1");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("horizontal", "general");
            _context.NewXmlWriter.WriteAttributeString("vertical", "bottom");
            _context.NewXmlWriter.WriteAttributeString("textRotation", "0");
            _context.NewXmlWriter.WriteAttributeString("wrapText", "0");
            _context.NewXmlWriter.WriteAttributeString("indent", "0");
            _context.NewXmlWriter.WriteAttributeString("relativeIndent", "0");
            _context.NewXmlWriter.WriteAttributeString("justifyLastLine", "0");
            _context.NewXmlWriter.WriteAttributeString("shrinkToFit", "0");
            _context.NewXmlWriter.WriteAttributeString("readingOrder", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("locked", "1");
            _context.NewXmlWriter.WriteAttributeString("hidden", "0");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
             *     <x:alignment horizontal="fill" />
             * </x:xf>
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "0");
            _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 0}");
            _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
            _context.NewXmlWriter.WriteAttributeString("xfId", "0");
            _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
            _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("horizontal", "fill");
            _context.NewXmlWriter.WriteEndElement();
            _context.NewXmlWriter.WriteEndElement();

            const int numFmtIndex = 166;
            var index = 0;
            foreach (var item in _context.ColumnsToApply)
            {
                index++;

                /*
                 * <x:xf numFmtId=""{numFmtIndex + i}"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
                 *     <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
                 *     <x:protection locked=""1"" hidden=""0"" />
                 * </x:xf>
                 */
                _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
                _context.NewXmlWriter.WriteAttributeString("numFmtId", (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString());
                _context.NewXmlWriter.WriteAttributeString("fontId", $"{_context.OldElementInfos.FontCount + 0}");
                _context.NewXmlWriter.WriteAttributeString("fillId", $"{_context.OldElementInfos.FillCount + 0}");
                _context.NewXmlWriter.WriteAttributeString("borderId", $"{_context.OldElementInfos.BorderCount + 1}");
                _context.NewXmlWriter.WriteAttributeString("xfId", "0");
                _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
                _context.NewXmlWriter.WriteAttributeString("applyFill", "1");
                _context.NewXmlWriter.WriteAttributeString("applyBorder", "1");
                _context.NewXmlWriter.WriteAttributeString("applyAlignment", "1");
                _context.NewXmlWriter.WriteAttributeString("applyProtection", "1");
                _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
                _context.NewXmlWriter.WriteAttributeString("horizontal", "general");
                _context.NewXmlWriter.WriteAttributeString("vertical", "bottom");
                _context.NewXmlWriter.WriteAttributeString("textRotation", "0");
                _context.NewXmlWriter.WriteAttributeString("wrapText", "0");
                _context.NewXmlWriter.WriteAttributeString("indent", "0");
                _context.NewXmlWriter.WriteAttributeString("relativeIndent", "0");
                _context.NewXmlWriter.WriteAttributeString("justifyLastLine", "0");
                _context.NewXmlWriter.WriteAttributeString("shrinkToFit", "0");
                _context.NewXmlWriter.WriteAttributeString("readingOrder", "0");
                _context.NewXmlWriter.WriteEndElement();
                _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
                _context.NewXmlWriter.WriteAttributeString("locked", "1");
                _context.NewXmlWriter.WriteAttributeString("hidden", "0");
                _context.NewXmlWriter.WriteEndElement();
                _context.NewXmlWriter.WriteEndElement();
            }
        }

        protected override async Task GenerateCellXfAsync()
        {
            /*
             * <x:xf></x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:alignment horizontal="left" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 2}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "left");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "general");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
             *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
             *     <x:protection locked="1" hidden="0" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "general");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            /*
             * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
             *     <x:alignment horizontal="fill" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "fill");
            await _context.NewXmlWriter.WriteEndElementAsync();
            await _context.NewXmlWriter.WriteEndElementAsync();

            const int numFmtIndex = 166;
            var index = 0;
            foreach (var item in _context.ColumnsToApply)
            {
                index++;

                /*
                 * <x:xf numFmtId=""{numFmtIndex + i}"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
                 *     <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
                 *     <x:protection locked=""1"" hidden=""0"" />
                 * </x:xf>
                 */
                await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString());
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1");
                await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI);
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "general");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0");
                await _context.NewXmlWriter.WriteEndElementAsync();
                await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI);
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1");
                await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0");
                await _context.NewXmlWriter.WriteEndElementAsync();
                await _context.NewXmlWriter.WriteEndElementAsync();
            }
        }
    }
}
