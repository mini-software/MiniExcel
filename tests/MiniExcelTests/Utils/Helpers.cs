namespace MiniExcelLibs.Tests.Utils
{
    using MiniExcelLibs.OpenXml;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Text;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;

    internal static class Helpers
    {
        private const int GENERAL_COLUMN_INDEX = 255;
        private const int MAX_COLUMN_INDEX = 16383;
        private static Dictionary<int, string> _IntMappingAlphabet;
        private static Dictionary<string, int> _AlphabetMappingInt;
        static Helpers()
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


        internal static string GetFirstSheetDimensionRefValue(string path)
        {
            var ns = new XmlNamespaceManager(new NameTable());
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            string refV;
            using (var stream = File.OpenRead(path))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
            {
                var sheet = archive.Entries.Single(w => w.FullName.StartsWith("xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase)
                    || w.FullName.StartsWith("/xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase)
                );
                using (var sheetStream = sheet.Open())
                {
                    var doc = XDocument.Load(sheetStream); ;
                    var dimension = doc.XPathSelectElement("/x:worksheet/x:dimension", ns);
                    refV = dimension.Attribute("ref").Value;
                }
            }

            return refV;
        }

    }

}
