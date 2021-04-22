namespace MiniExcelLibs.Utils
{
    using System;
    using System.Linq;
    using System.Text;
    using System.Xml;

    internal static class StringHelper
    {

        public static string GetLetter(string content)
        {
            //TODO:need to chekc
            return new String(content.Where(Char.IsLetter).ToArray());
        }

        public static int GetNumber(string content)
        {
            return int.Parse(new String(content.Where(Char.IsNumber).ToArray()));
        }

        /// <summary>
        /// Copy&Modify from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
        /// </summary>
        public static string ReadStringItem(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (reader.IsStartElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (reader.IsStartElement("r", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                {
                    result.Append(ReadRichTextRun(reader));
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result.ToString();
        }

        /// <summary>
        /// Copy&Modify from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
        /// </summary>
        private static string ReadRichTextRun(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (reader.IsStartElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                {
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return result.ToString();
        }
    }

}
