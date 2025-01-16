using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml.Styles
{
    internal class MinimalSheetStyleBuilder : SheetStyleBuilderBase
    {
        internal static SheetStyleElementInfos GenerateElementInfos = new SheetStyleElementInfos
        {
            NumFmtCount = 0,//默认的NumFmt数量是0，但是会有根据ColumnsToApply动态生成的NumFmt
            FontCount = 1,
            FillCount = 1,
            BorderCount = 1,
            CellStyleXfCount = 1,
            CellXfCount = 5
        };

        private readonly SheetStyleBuildContext _context;

        public MinimalSheetStyleBuilder(SheetStyleBuildContext context) : base(context)
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
                 * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
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
                 * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
                 */
                await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI);
                await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "numFmtId", _context.OldXmlReader.NamespaceURI, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString());
                await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "formatCode", _context.OldXmlReader.NamespaceURI, item.Format);
                await _context.NewXmlWriter.WriteFullEndElementAsync();
            }
        }

        protected override void GenerateFont()
        {
            /*
             * <x:font />
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
        }

        protected override async Task GenerateFontAsync()
        {
            /*
             * <x:font />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }

        protected override void GenerateFill()
        {
            /*
             * <x:fill />
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
        }

        protected override async Task GenerateFillAsync()
        {
            /*
             * <x:fill />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }

        protected override void GenerateBorder()
        {
            /*
             * <x:border />
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
        }

        protected override async Task GenerateBorderAsync()
        {
            /*
             * <x:border />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }

        protected override void GenerateCellStyleXf()
        {
            /*
             * <x:xf />
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
        }

        protected override async Task GenerateCellStyleXfAsync()
        {
            /*
             * <x:xf />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }

        protected override void GenerateCellXf()
        {
            /*
             * <x:xf />
             * <x:xf />
             * <x:xf />
             * <x:xf numFmtId="14" applyNumberFormat="1" />
             * <x:xf />
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", "14");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteFullEndElement();
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteFullEndElement();

            const int numFmtIndex = 166;
            var index = 0;
            foreach (var item in _context.ColumnsToApply)
            {
                index++;

                /*
                 * <x:xf numFmtId="{numFmtIndex + i}" applyNumberFormat="1" 
                 */
                _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
                _context.NewXmlWriter.WriteAttributeString("numFmtId", (numFmtIndex + index).ToString());
                _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
                _context.NewXmlWriter.WriteFullEndElement();
            }
        }

        protected override async Task GenerateCellXfAsync()
        {
            /*
             * <x:xf />
             * <x:xf />
             * <x:xf />
             * <x:xf numFmtId="14" applyNumberFormat="1" />
             * <x:xf />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "numFmtId", _context.OldXmlReader.NamespaceURI, "14");
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "applyNumberFormat", _context.OldXmlReader.NamespaceURI, "1");
            await _context.NewXmlWriter.WriteFullEndElementAsync();
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteFullEndElementAsync();

            const int numFmtIndex = 166;
            var index = 0;
            foreach (var item in _context.ColumnsToApply)
            {
                index++;

                /*
                 * <x:xf numFmtId="{numFmtIndex + i}" applyNumberFormat="1" 
                 */
                await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
                await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "numFmtId", _context.OldXmlReader.NamespaceURI, (numFmtIndex + index).ToString());
                await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "applyNumberFormat", _context.OldXmlReader.NamespaceURI, "1");
                await _context.NewXmlWriter.WriteFullEndElementAsync();
            }
        }
    }
}
