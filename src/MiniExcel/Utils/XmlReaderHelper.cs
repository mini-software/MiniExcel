using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MiniExcelLibs.Utils
{
    using System.Xml;

    internal static class XmlReaderHelper
    {
        /// <summary>
        /// Pass <?xml> and <worksheet>
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
                if (!XmlReaderHelper.SkipContent(reader))
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
            foreach (var ns in nss)
            {
                var attribute = reader.GetAttribute(name, ns);
                if (attribute != null)
                {
                    return attribute;
                }
            }

            return null;
        }
        
        public static IEnumerable<string> GetSharedStrings(Stream stream, params string[] nss)
        {
            using (var reader = XmlReader.Create(stream))
            {
                if (!XmlReaderHelper.IsStartElement(reader, "sst", nss))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "si", nss))
                    {
                        var value = StringHelper.ReadStringItem(reader);
                        yield return value;
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }
    }

}
