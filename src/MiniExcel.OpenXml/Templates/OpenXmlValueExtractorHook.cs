using System.ComponentModel;

namespace MiniExcelLib.OpenXml.Templates;



public static class OpenXmlValueExtractorHook
{

#if !NET8_0_OR_GREATER
    public class ReferenceComparer : IEqualityComparer<object>
    {
        public new bool Equals(object x, object y) => ReferenceEquals(x, y);

        public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
    }
#endif

    /// <summary>默认最大递归深度（可根据实际业务调整，10~15 通常足够）</summary>
    private const int DefaultMaxDepth = 4;

    /// <summary>
    /// 递归展开成 key.subkey 并完整格式化（防循环引用 / 防深度溢出）
    /// </summary>
    /// <param name="replacements">结果字典</param>
    /// <param name="key">当前键前缀</param>
    /// <param name="value">当前值</param>
    /// <param name="propInfo">当前属性信息（用于格式化特性读取）</param>
    /// <param name="maxDepth">最大递归深度，超出后将安全降级为 ToString()</param>
    public static void AddFlattenedAndFormattedValues(
        Dictionary<string, string> replacements,
        string key,
        object? value,
        PropertyInfo? propInfo = null,
        int maxDepth = DefaultMaxDepth)
    {
        // 使用引用相等比较器，防止业务类型重写 Equals/GetHashCode 导致误判
#if NET8_0_OR_GREATER
        var visited = new HashSet<object>(ReferenceEqualityComparer.Instance); 
#else
        var visited = new HashSet<object>(new ReferenceComparer());
#endif
        Core(replacements, key, value, propInfo, maxDepth, 0, visited);
    }

    private static void Core(
        Dictionary<string, string> replacements,
        string key,
        object? value,
        PropertyInfo? propInfo,
        int maxDepth,
        int currentDepth,
        HashSet<object> visited)
    {
        if (value == null || value.GetType() is not Type type)
        {
            replacements[key] = string.Empty;
            return;
        }

        // 1. 基础类型/枚举：直接格式化，不消耗深度，不进入引用追踪
        if (IsSimpleType(type) || type.IsEnum)
        {
            replacements[key] = GetFormattedValue(propInfo, value, type);
            return;
        }

        // 2. 深度控制：超出限制时安全降级，避免 OOM/StackOverflow
        if (currentDepth >= maxDepth)
        {
            replacements[key] = GetFormattedValue(propInfo, value, type);
            return;
        }

        // 3. 循环引用检测（仅针对引用类型，值类型不可能形成引用环）
        if (!type.IsValueType && !visited.Add(value))
        {
            replacements[key] = "[CircularReference]";
            return;
        }

        try
        {
            // 4. 字典处理
            if (value is Dictionary<string, object> dict)
            {
                foreach (var kv in dict)
                {
                    var subKey = string.Concat(key, ".", kv.Key);
                    Core(replacements, subKey, kv.Value, propInfo, maxDepth, currentDepth + 1, visited);
                }
                return;
            }

            // 5. 对象属性递归
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
               .Where(p => p.CanRead && p.GetIndexParameters().Length == 0);

            foreach (var prop in properties)
            {
                var subKey = string.Concat(key, ".", prop.Name);
                var subValue = prop.GetValue(value);
                Core(replacements, subKey, subValue, prop, maxDepth, currentDepth + 1, visited);
            }
        }
        finally
        {
            // 🔑 关键回溯：移除当前节点，允许同一对象在不同分支中被正常访问（DAG共享引用）
            // 仅拦截真正的“环”，不误杀合法的对象复用
            if (!type.IsValueType)
                visited.Remove(value);
        }
    }

    #region 你的完整格式化逻辑（独立维护，未改动核心行为）
    private static string GetFormattedValue(PropertyInfo? propInfo, object? cellValue, Type type)
    {
        string? cellValueStr;

        if (type == typeof(bool))
        {
            cellValueStr = (bool)cellValue! ? "1" : "0";
        }
        else if (type == typeof(DateTime))
        {
            cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
        }
        else if (type.IsEnum is true)
        {
            var stringValue = Enum.GetName(type, cellValue!) ?? "";
            var attr = type.GetField(stringValue)?.GetCustomAttribute<DescriptionAttribute>();
            var description = attr?.Description ?? stringValue;
            cellValueStr = XmlHelper.EncodeXml(description);
        }
        else
        {
            cellValueStr = XmlHelper.EncodeXml(cellValue?.ToString());
            if (TypeHelper.IsNumericType(type))
            {
                if (decimal.TryParse(cellValueStr, out var decimalValue))
                    cellValueStr = decimalValue.ToString(CultureInfo.InvariantCulture);
            }
        }

        var tempReplacement = cellValueStr ?? "";
        return tempReplacement.StartsWith("$=") || tempReplacement.StartsWith("=")
            ? $"&apos;{tempReplacement}"
            : tempReplacement;
    }

    private static bool IsSimpleType(Type type)
    {
        return type.IsPrimitive
            || type == typeof(string)
            || type == typeof(decimal)
            || type == typeof(DateTime)
            || type == typeof(Guid)
            || Nullable.GetUnderlyingType(type) != null;
    }

    private static string ConvertToDateTimeString(PropertyInfo? propInfo, object? cellValue)
    {
        if (propInfo == null || cellValue is not DateTime dt)
            return XmlHelper.EncodeXml(cellValue?.ToString() ?? "");

        return XmlHelper.EncodeXml(dt.ToString("yyyy-MM-dd HH:mm:ss"));
    }
    #endregion
}