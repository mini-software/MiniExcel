using MiniExcelLibs.OpenXml;
using System;
using System.Linq;
using System.Text;
using System.Xml;
    
namespace MiniExcelLibs.Utils
{
    internal static class StringHelper
    {
        private static readonly string[] _ns = { Config.SpreadsheetmlXmlns, Config.SpreadsheetmlXmlStrictns };
        //public static string GetLetter(string content) => content.FirstOrDefault(char.IsLetter).ToString();
        public static int GetDigit(string content) => content.FirstOrDefault(char.IsDigit) - '0';
        //public static int GetNumber(string content) => int.Parse(new string(content.Where(char.IsNumber).ToArray()));
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
        /// Copied and modified from ExcelDataReader - @MIT License
        /// </summary>
        public static string ReadStringItem(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", _ns))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result.Append(reader.ReadElementContentAsString());
                }
                else if (XmlReaderHelper.IsStartElement(reader, "r", _ns))
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
        /// Copied and modified from ExcelDataReader - @MIT License
        /// </summary>
        private static string ReadRichTextRun(XmlReader reader)
        {
            var result = new StringBuilder();
            if (!XmlReaderHelper.ReadFirstContent(reader))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", _ns))
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
