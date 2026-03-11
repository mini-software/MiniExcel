using MiniExcelLibs.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace MiniExcelLibs.OpenXml
{
    public sealed class ExcelColumnWidth
    {
        // Aptos is the default font for Office 2023 and onwards, over which the width of cells are calculated at the size of 11pt.
        // Priorly it was Calibri, which had very similar parameters, so no visual differences should be noticed.
        private const double DefaultCellPadding = 5;
        private const double Aptos11MaxDigitWidth = 7;
        public const double Aptos11Padding = DefaultCellPadding /  Aptos11MaxDigitWidth;
        
        public int Index { get; set; }
        public double Width { get; set; }
        public bool Hidden { get; set; }
        
        public static double GetWidthFromTextLength(double characters)
            => Math.Round(characters + Aptos11Padding, 8);
    }

    public sealed class ExcelWidthCollection : IReadOnlyCollection<ExcelColumnWidth>
    {
        private readonly Dictionary<int, ExcelColumnWidth> _columnWidths;
        private readonly double _maxWidth;

        public IReadOnlyCollection<ExcelColumnWidth> Columns
#if NET45
            => _columnWidths.Values.ToList();
#else
            => _columnWidths.Values;
#endif

        private ExcelWidthCollection(ICollection<ExcelColumnWidth> columnWidths, double maxWidth)
        {
            _maxWidth = ExcelColumnWidth.GetWidthFromTextLength(maxWidth);
            _columnWidths = columnWidths.ToDictionary(x => x.Index);
        }

        internal static ExcelWidthCollection FromProps(ICollection<ExcelColumnInfo> mappings, double? minWidth = null, double maxWidth = 200)
        {
            var i = 1;
            var columnWidths = new List<ExcelColumnWidth>();

            foreach (var map in mappings)
            {
                if ((map?.ExcelHidden ?? false) || map?.ExcelColumnWidth != null || minWidth != null)
                {
                    var colIndex = map?.ExcelColumnIndex + 1 ?? i;
                    var hidden = map?.ExcelHidden ?? false;
                    var width = map?.ExcelColumnWidth ?? minWidth ?? 8.42857143;

                    columnWidths.Add(new ExcelColumnWidth { Index = colIndex, Width = width + ExcelColumnWidth.Aptos11Padding, Hidden = hidden});
                }

                i++;
            }

            return new ExcelWidthCollection(columnWidths, maxWidth);
        }

        internal void AdjustWidth(int columnIndex, string columnValue)
        {
            if (!string.IsNullOrEmpty(columnValue) && _columnWidths.TryGetValue(columnIndex, out var currentWidth))
            {
                var desiredWidth = ExcelColumnWidth.GetWidthFromTextLength(columnValue.Length);
                var adjustedWidth = Math.Max(currentWidth.Width, desiredWidth);
                currentWidth.Width = Math.Min(_maxWidth, adjustedWidth);
            }
        }

        public IEnumerator<ExcelColumnWidth> GetEnumerator() => _columnWidths.Values.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public int Count =>  _columnWidths.Count;
    }
}
