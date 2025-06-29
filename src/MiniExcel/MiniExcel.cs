using System.Dynamic;
using MiniExcelLib.Core.OpenXml;
using MiniExcelLib.Core.OpenXml.Templates;

namespace MiniExcelLib;

public static partial class MiniExcel
{
    public static partial class Exporter;
    
    public static partial class Importer
    {
        private static IDictionary<string, object?> GetNewExpandoObject() => new ExpandoObject();
        private static IDictionary<string, object?> AddPairToDict(IDictionary<string, object?> dict, KeyValuePair<string, object?> pair)
        {
            dict.Add(pair);
            return dict; 
        }
    }

    public static partial class Templater
    {
        private static OpenXmlTemplate GetOpenXmlTemplate(Stream stream, OpenXmlConfiguration? configuration)
        {
            var valueExtractor = new OpenXmlValueExtractor();
            return new OpenXmlTemplate(stream, configuration, valueExtractor);
        }
    }
}