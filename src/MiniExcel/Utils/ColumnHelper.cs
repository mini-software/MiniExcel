namespace MiniExcelLibs.Utils
{
    using System.Collections.Generic;
    using System.IO;

    // For Row/Column Index
    internal static partial class ColumnHelper
    {
        private const int GENERAL_COLUMN_INDEX = 255;
        private const int MAX_COLUMN_INDEX = 16383;
        private static Dictionary<int, string> _IntMappingAlphabet;
        private static Dictionary<string, int> _AlphabetMappingInt;
        static ColumnHelper()
        {
            if (_IntMappingAlphabet == null && _AlphabetMappingInt == null)
            {
                _IntMappingAlphabet = new Dictionary<int, string>();
                _AlphabetMappingInt = new Dictionary<string, int>();
                for (int i = 0; i <= GENERAL_COLUMN_INDEX; i++)
                {
                    _IntMappingAlphabet.Add(i, IntToLetters(i));
                    _AlphabetMappingInt.Add(IntToLetters(i), i);
                }
            }
        }

        public static string GetAlphabetColumnName(int columnIndex)
        {
            CheckAndSetMaxColumnIndex(columnIndex);
            return _IntMappingAlphabet[columnIndex];
        }

        public static int GetColumnIndex(string columnName)
        {
            var columnIndex = _AlphabetMappingInt[columnName];
            CheckAndSetMaxColumnIndex(columnIndex);
            return columnIndex;
        }

        private static void CheckAndSetMaxColumnIndex(int columnIndex)
        {
            if (columnIndex >= _IntMappingAlphabet.Count)
            {
                if (columnIndex > MAX_COLUMN_INDEX)
                    throw new InvalidDataException($"ColumnIndex {columnIndex} over excel vaild max index.");
                for (int i = _IntMappingAlphabet.Count; i <= columnIndex; i++)
                {
                    _IntMappingAlphabet.Add(i, IntToLetters(i));
                    _AlphabetMappingInt.Add(IntToLetters(i), i);
                }
            }
        }

        internal static string IntToLetters(int value)
        {
            value = value + 1;
            string result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }
    }

}
