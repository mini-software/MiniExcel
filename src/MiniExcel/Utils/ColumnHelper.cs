using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs.Utils
{

    // For Row/Column Index
    internal static class ColumnHelper
    {
        private const int GENERAL_COLUMN_INDEX = 255;
        private const int MAX_COLUMN_INDEX = 16383;
        private static int _IntMappingAlphabetCount = 0;
        private static readonly ConcurrentDictionary<int, string> _IntMappingAlphabet = new ConcurrentDictionary<int, string>();
        private static readonly ConcurrentDictionary<string, int> _AlphabetMappingInt = new ConcurrentDictionary<string, int>();

        static ColumnHelper()
        {
            _IntMappingAlphabetCount = _IntMappingAlphabet.Count;
            CheckAndSetMaxColumnIndex(GENERAL_COLUMN_INDEX);
        }

        public static string GetAlphabetColumnName(int columnIndex)
        {
            CheckAndSetMaxColumnIndex(columnIndex);
            return _IntMappingAlphabet.TryGetValue(columnIndex, out var value) ? value
                : throw new KeyNotFoundException();
        }

        public static int GetColumnIndex(string columnName)
        {
            if (_AlphabetMappingInt.TryGetValue(columnName, out var columnIndex))
                CheckAndSetMaxColumnIndex(columnIndex);
            return columnIndex;
        }

        private static void CheckAndSetMaxColumnIndex(int columnIndex)
        {
            if (columnIndex < _IntMappingAlphabetCount)
                return;
            if (columnIndex > MAX_COLUMN_INDEX)
                throw new InvalidDataException($"ColumnIndex {columnIndex} over excel vaild max index.");
            
            for (int i = _IntMappingAlphabet.Count; i <= columnIndex; i++)
            {
                _IntMappingAlphabet.AddOrUpdate(i, IntToLetters(i), (a, b) => IntToLetters(i));
                _AlphabetMappingInt.AddOrUpdate(IntToLetters(i), i, (a, b) => i);
            }
            _IntMappingAlphabetCount = _IntMappingAlphabet.Count;
        }

        internal static string IntToLetters(int value)
        {
            value++;
            var result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }
    }
}
