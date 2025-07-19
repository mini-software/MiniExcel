namespace MiniExcelLib.Helpers;

/// <summary> XmlEncoder MIT Copyright ©2021 from https://github.com/ClosedXML </summary>
internal static partial class XmlHelper
{
#if NET7_0_OR_GREATER
    [GeneratedRegex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled)] private static partial Regex X4LRegexImpl();
    private static readonly Regex X4LRegex = X4LRegexImpl();
    
    [GeneratedRegex("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled)] private static partial Regex UppercaseX4LRegexImpl();
    private static readonly Regex UppercaseX4LRegex = UppercaseX4LRegexImpl();
#else
    private static readonly Regex X4LRegex = new("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);
    private static readonly Regex UppercaseX4LRegex = new("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled);
#endif

    public static StringBuilder? EncodeString(string? encodeStr)
    {
        if (encodeStr is null)
            return null;

        encodeStr = X4LRegex.Replace(encodeStr, "_x005F_$1_");

        var sb = new StringBuilder(encodeStr.Length);
        foreach (var ch in encodeStr)
        {
            if (XmlConvert.IsXmlChar(ch))
                sb.Append(ch);
            else
                sb.Append(XmlConvert.EncodeName(ch.ToString()));
        }

        return sb;
    }

    public static string? DecodeString(string? decodeStr)
    {
        if (string.IsNullOrEmpty(decodeStr))
            return string.Empty;

        decodeStr = UppercaseX4LRegex.Replace(decodeStr, "_x005F_$1_");
        return XmlConvert.DecodeName(decodeStr);
    }
    
    public static string MinifyXml(string xml) => xml
        .Replace("\r", "")
        .Replace("\n", "")
        .Replace("\t", "")
        .Trim();

    /// <summary>
    /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
    /// </summary>
    public static string EncodeXml(string? value) => value is null ? ""
        : XmlHelper.EncodeString(value)
            ?.Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;")
            .ToString() ?? "";
}