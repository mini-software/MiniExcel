using System.Resources;

namespace MiniExcelLib.Core.Attributes;

public class MiniExcelColumnNameAttribute(string columnName, string[]? aliases = null) : MiniExcelAttributeBase
{
    private ResourceManager? _resourceManager;

    [Obsolete("Please use the \"Name\" property instead")]
    public string ExcelColumnName { get; set; } = columnName;
    public string Name { get; set; } = columnName;
    public string[] Aliases { get; set; } = aliases ?? [];
    
    private Type? _resourceType;
    public Type? ResourceType
    {
        get => _resourceType;
        set
        {
            if (_resourceType == value) 
                return;

            _resourceType = value;
            if (value is null) 
                return;

            const BindingFlags bindingFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
            if (value.GetProperty(nameof(ResourceManager), bindingFlags) is { } property && 
                property.GetValue(null) is ResourceManager resourceManager)
            {
                _resourceManager = resourceManager;
            }
            else
            {
                _resourceManager = new ResourceManager(value);
            }
        }
    }

    internal string? GetColumnName() => !string.IsNullOrEmpty(Name) 
        ? _resourceManager?.GetString(Name) ?? Name 
        : null;
}
