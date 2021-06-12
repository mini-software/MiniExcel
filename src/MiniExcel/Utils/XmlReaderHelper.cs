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
    }

}
