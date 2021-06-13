namespace MiniExcelLibs.Utils
{
    using MiniExcelLibs.OpenXml;
    using System;
    using System.Linq;
    using System.Text;
    using System.Xml;

    internal static class StringHelper
    {
        private const string _ns = Config.SpreadsheetmlXmlns;
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
        /// Copy&Modify from ExcelDataReader @MIT License
        /// </summary>
        public static string ReadStringItem(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (reader.IsStartElement("t", _ns))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (reader.IsStartElement("r", _ns))
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
        /// Copy&Modify from ExcelDataReader @MIT License
        /// </summary>
        private static string ReadRichTextRun(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (reader.IsStartElement("t", _ns))
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
