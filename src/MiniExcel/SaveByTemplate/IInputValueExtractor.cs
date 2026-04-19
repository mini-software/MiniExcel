namespace MiniExcelLibs.OpenXml.SaveByTemplate;

public interface IInputValueExtractor
{
    IDictionary<string, object> ToValueDictionary(object valueObject);
}