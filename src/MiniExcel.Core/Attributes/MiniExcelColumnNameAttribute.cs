using System.Resources;

namespace MiniExcelLib.Core.Attributes;

public class MiniExcelColumnNameAttribute(string columnName, string[]? aliases = null) : MiniExcelAttributeBase
{
    public string Name { get; set; } = columnName;
    public string[] Aliases { get; set; } = aliases ?? [];
    
    private ResourceManager? _resourceManager;
    public Type? ResourceType
    {
        get;
        set
        {
            if (field == value) 
                return;

            field = value;
            if (value is null) 
                return;

            const BindingFlags bindingFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
            _resourceManager = value.GetProperty(nameof(ResourceManager), bindingFlags) is { } property 
                ? property.GetValue(null) as ResourceManager
                : new ResourceManager(value);
        }
    }

    internal string GetColumnName() => !string.IsNullOrEmpty(Name) 
        ? _resourceManager?.GetString(Name) ?? Name 
        : "";
}
