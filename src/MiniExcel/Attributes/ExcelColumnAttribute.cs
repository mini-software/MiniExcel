using MiniExcelLibs.Utils;
using System;

namespace MiniExcelLibs.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        private int _index = -1;
        private string _xName;

        internal int FormatId { get; set; } = -1;

        public string Name { get; set; }

        public string[] Aliases { get; set; }

        public double Width { get; set; } = 9.28515625;

        public string Format { get; set; }

        public bool Ignore { get; set; }

        public int Index
        {
            get => _index;
            set => Init(value);
        }

        public string IndexName
        {
            get => _xName;
            set => Init(ColumnHelper.GetColumnIndex(value), value);
        }

        private void Init(int index, string columnName = null)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index), index,
                    $"Column index {index} must be greater or equal to zero.");
            }

            if (_xName == null)
                if (columnName != null)
                    _xName = columnName;
                else
                    _xName = ColumnHelper.GetAlphabetColumnName(index);
            _index = index;
        }
    }


    public class DynamicExcelColumn : ExcelColumnAttribute
    {
        public string Key { get; set; }

        public DynamicExcelColumn(string key)
        {
            Key = key;
        }
    }
}