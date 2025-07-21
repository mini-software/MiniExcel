namespace MiniExcelLib.Core.Abstractions;

public interface IInputValueExtractor
{
    IDictionary<string, object?> ToValueDictionary(object valueObject);
}