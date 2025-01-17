namespace MiniExcelLibs
{
    using System;
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
        private bool _disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="MiniExcelDataReader"/> class.
        /// </summary>
        /// <param name="stream">The stream to read from.</param>
        /// <param name="useHeaderRow">Whether to use the header row.</param>
        /// <param name="sheetName">The name of the sheet.</param>
        /// <param name="excelType">The type of the Excel file.</param>
        /// <param name="startCell">The start cell.</param>
        /// <param name="configuration">The configuration.</param>
        internal MiniExcelDataReader(Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _source = MiniExcel.Query(_stream, useHeaderRow, sheetName, excelType, startCell, configuration).Cast<IDictionary<string, object>>().GetEnumerator();
            if (_source.MoveNext())
            {
                _keys = _source.Current?.Keys.ToList() ?? new List<string>();
                _fieldCount = _keys.Count;
            }
        }

        /// <inheritdoc/>
        public override object GetValue(int i)
        {
            if (_source.Current == null)
                throw new InvalidOperationException("No current row available.");
            return _source.Current[_keys[i]];
        }

        /// <inheritdoc/>
        public override int FieldCount => _fieldCount;

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

        /// <inheritdoc/>
        public override string GetName(int i)
        {
            return _keys[i];
        }

        /// <inheritdoc/>
        public override int GetOrdinal(string name)
        {
            _keys.IndexOf(name);
            return _keys.IndexOf(name);
        }

        /// <inheritdoc/>
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _stream?.Dispose();
                }
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Disposes the object.
        /// </summary>
        public new void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
