using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MiniExcelLibs.OpenXml
{
    public sealed class ExcelColumnWidth
    {
        public int Index { get; set; }
        public double Width { get; set; }

        internal static IEnumerable<ExcelColumnWidth> FromProps(IEnumerable<ExcelColumnInfo> props, double? minWidth = null)
        {
            var i = 1;
            foreach (var p in props)
            {
                if (p == null || (p.ExcelColumnWidth == null && minWidth == null))
                {
                    i++;
                    continue;
                }
                var colIndex = p.ExcelColumnIndex == null ? i : p.ExcelColumnIndex.GetValueOrDefault() + 1;
                yield return new ExcelColumnWidth
                {
                    Index = colIndex,
                    Width = p.ExcelColumnWidth ?? minWidth.Value,
                };
                i++;
            }
        }
    }

    public sealed class ExcelWidthCollection
    {
        private readonly Dictionary<int, ExcelColumnWidth> _columnWidths;
        private readonly double _maxWidth;

        public IEnumerable<ExcelColumnWidth> Columns => _columnWidths.Values;

        internal ExcelWidthCollection(double minWidth, double maxWidth, IEnumerable<ExcelColumnInfo> props)
        {
            _maxWidth = maxWidth;
            _columnWidths = ExcelColumnWidth.FromProps(props, minWidth).ToDictionary(x => x.Index);
        }

        public void AdjustWidth(int columnIndex, string columnValue)
        {
            if (string.IsNullOrEmpty(columnValue) || !_columnWidths.TryGetValue(columnIndex, out var currentWidth))
            {
                return;
            }

            var adjustedWidth = Math.Max(currentWidth.Width, GetApproximateRequiredCalibriWidth(columnValue.Length));
            currentWidth.Width = Math.Min(_maxWidth, adjustedWidth);
        }

        /// <summary>
        /// Get the approximate width of the given text for Calibri 11pt
        /// </summary>
        /// <remarks>
        /// Rounds the result to 2 decimal places.
        /// </remarks>
        public static double GetApproximateRequiredCalibriWidth(int textLength)
        {
            double characterWidthFactor = 1.2;  // Estimated factor for Calibri, 11pt
            double padding = 2;  // Add some padding for extra spacing

            double excelColumnWidth = (textLength * characterWidthFactor) + padding;

            return Math.Round(excelColumnWidth, 2);
        }
    }
}
