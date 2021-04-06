namespace MiniExcelLibs.Utils
{
    using System.Globalization;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;

    /// <summary> XmlEncoder MIT Copyright ©2021 from https://github.com/ClosedXML </summary>
    internal static class XmlEncoder
    {
        private static readonly Regex xHHHHRegex = new Regex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);
        private static readonly Regex Uppercase_X_HHHHRegex = new Regex("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled);

        public static string EncodeString(string encodeStr)
        {
            if (encodeStr == null) return null;

            encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

            var sb = new StringBuilder(encodeStr.Length);

            foreach (var ch in encodeStr)
            {
                if (XmlConvert.IsXmlChar(ch))
                    sb.Append(ch);
                else
                    sb.Append(XmlConvert.EncodeName(ch.ToString()));
            }

            return sb.ToString();
        }

        public static string DecodeString(string decodeStr)
        {
            if (string.IsNullOrEmpty(decodeStr))
                return string.Empty;
            decodeStr = Uppercase_X_HHHHRegex.Replace(decodeStr, "_x005F_$1_");
            return XmlConvert.DecodeName(decodeStr);
        }


        private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");
        public static string ConvertEscapeChars(string input)
        {
            return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
        }
    }

}
