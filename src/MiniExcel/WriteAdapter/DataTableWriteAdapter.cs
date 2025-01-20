using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Data;
using System.Threading;

namespace MiniExcelLibs.WriteAdapter
{
    internal class DataTableWriteAdapter : IMiniExcelWriteAdapter
    {
        private readonly DataTable _dataTable;
        private readonly Configuration _configuration;

        public DataTableWriteAdapter(DataTable dataTable, Configuration configuration)
        {
            _dataTable = dataTable;
            _configuration = configuration;
        }

        public bool TryGetKnownCount(out int count)
        {
            count = _dataTable.Rows.Count;
            return true;
        }

        public List<ExcelColumnInfo> GetColumns()
        {
            var props = new List<ExcelColumnInfo>();
            for (var i = 0; i < _dataTable.Columns.Count; i++)
            {
                var columnName = _dataTable.Columns[i].Caption ?? _dataTable.Columns[i].ColumnName;
                var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                props.Add(prop);
            }
            return props;
        }

        public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<ExcelColumnInfo> props, CancellationToken cancellationToken = default)
        {
            for (int row = 0; row < _dataTable.Rows.Count; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValues(row, props);
            }
        }

        private IEnumerable<CellWriteInfo> GetRowValues(int row, List<ExcelColumnInfo> props)
        {
            for (int i = 0, column = 1; i < _dataTable.Columns.Count; i++, column++)
            {
                yield return new CellWriteInfo(_dataTable.Rows[row][i], column, props[i]);
            }
        }
    }
}


