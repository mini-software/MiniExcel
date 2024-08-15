namespace MiniExcelLibs
{
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;

    public class MiniExcelDataReader : MiniExcelDataReaderBase
    {
        private readonly IEnumerator<IDictionary<string, object>> _source;
        private readonly int _fieldCount;
        private readonly List<string> _keys;
        private readonly Stream _stream;
        private bool _isFirst = true;

        internal MiniExcelDataReader(Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            _stream = stream;
            _source = MiniExcel.Query(_stream, useHeaderRow, sheetName, excelType, startCell, configuration).Cast<IDictionary<string, object>>().GetEnumerator();
            var isNext = _source.MoveNext();
            if (isNext)
            {
                _keys = _source.Current.Keys.ToList();
                _fieldCount = _keys.Count;
            }
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public override object GetValue(int i)
        {
            return _source.Current[_keys[i]];
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public override int FieldCount
        {
            get { return _fieldCount; }
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        /// <returns></returns>
        public override bool Read()
        {
            if (_isFirst)
            {
                _isFirst = false;
                return true;
            }
            return _source.MoveNext();
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public override string GetName(int i)
        {
            return _keys[i];
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public override int GetOrdinal(string name)
        {
            var i = _keys.IndexOf(name);
            return _keys.IndexOf(name);
        }

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _stream.Dispose();
            }
        }
    }
}
