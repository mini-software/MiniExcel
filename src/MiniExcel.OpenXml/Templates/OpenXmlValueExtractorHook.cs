using System.ComponentModel;
using System.Xml.Linq;

namespace MiniExcelLib.OpenXml.Templates;



internal partial class OpenXmlTemplate
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

    #endregion


    private async Task<bool> HookSheetProcess(OpenXmlZip outputFileArchive, string realSheetName, IDictionary<int, string> templateSharedStrings, int sheetIdx, List<(int Index, string Name)> allSheetInfos, ZipArchiveEntry templateSheet, string templateFullName, IDictionary<string, object?> inputValues, CancellationToken cancellationToken)
    {
        var m = Regex.Match(realSheetName, @"\$([^$]+)\$");
        if (m.Success && inputValues.TryGetValue(m.Groups[1].Value, out var subObj) && subObj is IEnumerable sunIter)
        {
            // 基础表名（从模板占位符提取）
            var baseSheetName = m.Groups[1].Value;
            var subIndex = 1;

            // 1. 【先批量创建所有工作表文件】（流自动关闭，无冲突）
            foreach (var subRoot in sunIter)
            {
                _xRowInfos.Clear();
                _xMergeCellInfos.Clear();
                _newXMergeCellInfos.Clear();
                _calcChainCellRefs.Clear();

                var subValues = _inputValueExtractor.ToValueDictionary(subRoot);

                sheetIdx++;
                var newSheetPath = $"xl/worksheets/sheet{sheetIdx}.xml";

                // 处理表名
                string finalSheetName;
                if (subValues.TryGetValue("SheetName", out var customSheetName) && customSheetName != null)
                {
                    finalSheetName = customSheetName.ToString()?.Trim() ?? $"{baseSheetName}{subIndex++}";
                }
                else
                {
                    finalSheetName = $"{baseSheetName}{sheetIdx++}";
                }

                // 🔥 关键：只收集，不调用配置方法
                allSheetInfos.Add((sheetIdx, finalSheetName));


                // 创建工作表（独立作用域，流自动关闭）
                var newSheetEntry = outputFileArchive.ZipFile.CreateEntry(newSheetPath);
#if NET8_0_OR_GREATER
                    await using var newSheetStream = await newSheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
                using var newSheetStream = newSheetEntry.Open();
#endif
                await GenerateSheetByCreateModeAsync(templateSheet, newSheetStream, subValues, templateSharedStrings, cancellationToken: cancellationToken).ConfigureAwait(false);


                _calcChainContent.Append(CalcChainHelper.GetCalcChainContent(_calcChainCellRefs, sheetIdx));
            }

            return true;
        }

        return false;
    }


    /// <summary>
    /// 批量添加工作表到Excel配置（一次性写入，无覆盖，表名生效）
    /// </summary>
    [CreateSyncVersion]
    public static async Task BatchAddSheetsToExcelConfigAsync(ZipArchive outputZip, ZipArchive templateArchive, List<(int Index, string Name)> sheetInfos, CancellationToken cancellationToken)
    {
        // ======================================
        // 阶段1：纯内存读取并修改配置（读完立即关闭所有流）
        // ======================================
        XDocument relDoc = await LoadTemplateXmlAsync(templateArchive, "xl/_rels/workbook.xml.rels", cancellationToken).ConfigureAwait(false);
        XDocument wbDoc = await LoadTemplateXmlAsync(templateArchive, "xl/workbook.xml", cancellationToken).ConfigureAwait(false);

        // 命名空间
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // 1. 🗑️ 清空 workbook.xml 中的所有 <sheet>，重建干净的 <sheets> 容器
        var sheetsPart = wbDoc.Root?.Element(ssNs + "sheets");
        if (sheetsPart != null)
        {
            // 直接清空子节点，保留容器本身及默认命名空间（避免命名空间丢失导致 Excel 报错）
            sheetsPart.Elements().Remove();
        }
        else
        {
            // 若原模板无 sheets 节点，则新建一个追加到根节点
            wbDoc.Root?.Add(new XElement(ssNs + "sheets"));
        }

        // 2. 🔗 清理 workbook.xml.rels 中所有指向 worksheet 的关系记录
        const string worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

        var relsRoot = relDoc.Root;
        if (relsRoot != null)
        {
            // 仅删除 Type 为 worksheet 的关系，保留 sharedStrings/styles/theme 等核心关系
            var worksheetRels = relsRoot.Elements(relNs + "Relationship")
                .Where(r => r.Attribute("Type")?.Value == worksheetRelType)
                .ToList();

            foreach (var rel in worksheetRels) rel.Remove();
        }


        // 批量添加关系
        foreach (var sheet in sheetInfos)
        {
            relDoc.Root!.Add(new XElement(relNs + "Relationship",
                new XAttribute("Id", $"rIdSheet{sheet.Index}"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", $"worksheets/sheet{sheet.Index}.xml")));
        }

        // 批量添加工作表
        var sheetsNode = wbDoc.Descendants(ssNs + "sheets").FirstOrDefault();
        if (sheetsNode != null)
        {
            foreach (var sheet in sheetInfos)
            {
                sheetsNode.Add(new XElement(ssNs + "sheet",
                    new XAttribute("name", sheet.Name),
                    new XAttribute("sheetId", sheet.Index),
                    new XAttribute(rNs + "id", $"rIdSheet{sheet.Index}")));
            }
        }

        // ======================================
        // 阶段2：所有流已关闭 → 安全创建条目并写入
        // ======================================
        await SaveXmlToZipAsync(outputZip, "xl/_rels/workbook.xml.rels", relDoc, cancellationToken).ConfigureAwait(false);
        await SaveXmlToZipAsync(outputZip, "xl/workbook.xml", wbDoc, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 读取模板XML（流自动关闭，返回内存XDocument）
    /// </summary>
    [CreateSyncVersion]
    private static async Task<XDocument> LoadTemplateXmlAsync(ZipArchive templateArchive, string path, CancellationToken cancellationToken)
    {
        var entry = templateArchive.GetEntry(path)!;
#if NET8_0_OR_GREATER
    await using var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
    return await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var stream = entry.Open();
        return XDocument.Load(stream);
#endif
    }

    /// <summary>
    /// 写入XML到压缩包（流自动关闭，无残留）
    /// </summary>
    [CreateSyncVersion]
    private static async Task SaveXmlToZipAsync(ZipArchive outputZip, string path, XDocument doc, CancellationToken cancellationToken)
    {
        // 🔥 此时无任何打开流，安全创建条目


        var newEntry = outputZip.CreateEntry(path);
#if NET8_0_OR_GREATER
    await using var stream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
    await doc.SaveAsync(stream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var stream = newEntry.Open();
        doc.Save(stream);
#endif
    }




    /// <summary>
    /// 精准解析：获取【真实工作表名】+ 对应sheet xml路径
    /// </summary>
    public Dictionary<string, string> GetRealSheetNameMap(ZipArchive archive)
    {
        var nameToPath = new Dictionary<string, string>();
        var ridToSheetPath = new Dictionary<string, string>();

        // 1. 读取 workbook.xml.rels 拿到 rId → sheet文件路径
        var relsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/_rels/workbook.xml.rels");
        if (relsEntry == null) return nameToPath;

        using var relStream = relsEntry.Open();
        var relDoc = XDocument.Load(relStream);
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";

        foreach (var rel in relDoc.Descendants(relNs + "Relationship"))
        {
            var rid = rel.Attribute("Id")?.Value;
            var target = rel.Attribute("Target")?.Value;
            if (string.IsNullOrEmpty(rid) || string.IsNullOrEmpty(target)) continue;
            // 拼接完整路径
            var fullSheetPath = Path.Combine("xl", target).Replace("\\", "/");
            ridToSheetPath[rid] = fullSheetPath;
        }

        // 2. 读取 workbook.xml 拿到 真实表名 + rId
        var wbEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/workbook.xml");
        if (wbEntry == null) return nameToPath;

        using var wbStream = wbEntry.Open();
        var wbDoc = XDocument.Load(wbStream);
        XNamespace ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (var sheetNode in wbDoc.Descendants(ssNs + "sheet"))
        {
            var realName = sheetNode.Attribute("name")?.Value?.Trim() ?? "";
            var rid = sheetNode.Attribute(rNs + "id")?.Value ?? "";
            if (string.IsNullOrEmpty(realName) || string.IsNullOrEmpty(rid)) continue;

            if (ridToSheetPath.TryGetValue(rid, out var sheetPath))
            {
                nameToPath[sheetPath] = realName; // key:xml路径  value:真实表名
            }
        }
        return nameToPath;
    }





}