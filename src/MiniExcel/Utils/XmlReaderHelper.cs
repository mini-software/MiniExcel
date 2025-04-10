using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace MiniExcelLibs.Utils
{
    internal static class XmlReaderHelper
    {
        /// <summary>
        /// Pass &lt;?xml&gt; and &lt;worksheet&gt;
        /// </summary>
        /// <param name="reader"></param>
        public static void PassXmlDeclartionAndWorksheet(this XmlReader reader)
        {
            reader.MoveToContent();
            reader.Read();
        }

        /// <summary>
        /// e.g skip row 1 to row 2
        /// </summary>
        /// <param name="reader"></param>
        public static void SkipToNextSameLevelDom(XmlReader reader)
        {
            while (!reader.EOF)
            {
                if (!SkipContent(reader))
                    break;
            }
        }

        //Method from ExcelDataReader @MIT License
        public static bool ReadFirstContent(XmlReader reader)
        {
            if (reader.IsEmptyElement)
            {
                reader.Read();
                return false;
            }

            reader.MoveToContent();
            reader.Read();
            return true;
        }

        //Method from ExcelDataReader @MIT License
        public static bool SkipContent(XmlReader reader)
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                reader.Read();
                return false;
            }

            reader.Skip();
            return true;
        }

        public static bool IsStartElement(XmlReader reader, string name, params string[] nss)
        {
            return nss.Any(s => reader.IsStartElement(name, s));
        }

        public static string GetAttribute(XmlReader reader, string name, params string[] nss)
        {
            return nss
                .Select(ns => reader.GetAttribute(name, ns))
                .FirstOrDefault(at => at != null);
        }

        public static IEnumerable<string> GetSharedStrings(Stream stream, params string[] nss)
        {
            using (var reader = XmlReader.Create(stream))
            {
                if (!IsStartElement(reader, "sst", nss))
                    yield break;

                if (!ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (IsStartElement(reader, "si", nss))
                    {
                        var value = StringHelper.ReadStringItem(reader);
                        yield return value;
                    }
                    else if (!SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }
    }
}
