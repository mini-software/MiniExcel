using System.Collections.Generic;

namespace MiniExcelLibs.SaveByTemplate;

public interface IInputValueExtractor
{
    IDictionary<string, object> ToValueDictionary(object valueObject);
}