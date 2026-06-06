namespace MiniExcelLib.OpenXml.Templates;

internal partial class OpenXmlTemplate
{
    private static readonly XNamespace PackageRelNs = Schemas.OpenXmlPackageRelationships;
    private static readonly XNamespace SpreadsheetRelNs = Schemas.SpreadsheetmlXmlRelationships;

#if NET
    [GeneratedRegex(@"\$([^$]+)\$")] private static partial Regex ParametrizedSheetRegex();
    private static readonly Regex ParametrizedSheetRegexImpl = ParametrizedSheetRegex();
#else
    private static readonly Regex ParametrizedSheetRegexImpl = new(@"\$([^$]+)\$", RegexOptions.Compiled);
#endif

    /// <summary>
    /// Recursively flattens an object graph into a dictionary of "key.subkey" pairs and fully formats the values.
    /// Includes protection against circular references and stack overflow via depth limiting.
    /// </summary>
    private static void FlattenAndFormatValues(Dictionary<string, string> replacements, string key, object? value, int maxDepth, PropertyInfo? propInfo = null)
    {
        // Initialize a HashSet with reference equality comparer to track visited objects and prevent infinite loops from circular references.
        var visited = new HashSet<object>(ReferenceEqualityComparer.Instance); 

        // Start the recursive processing with initial depth set to 0.
        TraverseAndFlatten(replacements, key, value, propInfo, maxDepth, 0, visited);
        return;

        // <summary>
        // The internal recursive method that performs the actual object traversal, flattening, and formatting.
        // </summary>
        static void TraverseAndFlatten(
            Dictionary<string, string> replacements,
            string key,
            object? value,
            PropertyInfo? propInfo,
            int maxDepth,
            int currentDepth,
            HashSet<object> visited)
        {
            // Handle null values or invalid types
            if (value?.GetType() is not { } type)
            {
                replacements[key] = string.Empty;
                return;
            }

            // 1. Primitive types / Enums: Format directly, do not consume depth and do not enter reference tracking.
            if (type.IsPrimitive || type.IsEnum ||
                type == typeof(string) || type == typeof(decimal) ||
                type == typeof(DateTime) || type == typeof(Guid) ||
                Nullable.GetUnderlyingType(type) != null)
            {
                replacements[key] = GetFormattedValue(propInfo, value, type);
                return;
            }

            // 2. Depth control: Safe fallback to string representation when exceeding the limit to avoid OOM/StackOverflow.
            if (currentDepth >= maxDepth)
                return;

            // 3. Circular reference detection (only for reference types; value types cannot form reference loops).
            if (!type.IsValueType && !visited.Add(value))
                return;

            try
            {
                // 4. Dictionary handling: Iterate through key-value pairs and recursively process values.
                if (value is Dictionary<string, object> dict)
                {
                    foreach (var (innerKey, innerValue) in dict)
                    {
                        // Construct the new sub-key by appending the dictionary key.
                        var subKey = string.Concat(key, ".", innerKey);
                        TraverseAndFlatten(replacements, subKey, innerValue, propInfo, maxDepth, currentDepth + 1, visited);
                    }
                    return;
                }

                // 5. Object property recursion: Get public instance properties filtering out indexers and write-only properties.
                var properties = type
                    .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.CanRead && p.GetIndexParameters().Length == 0);

                foreach (var prop in properties)
                {
                    // Construct the new sub-key by appending the property name
                    var subKey = string.Concat(key, ".", prop.Name);
                    var subValue = prop.GetValue(value);
                    TraverseAndFlatten(replacements, subKey, subValue, prop, maxDepth, currentDepth + 1, visited);
                }
            }
            finally
            {
                // After loops (A -> B -> A) are excluded remove the current node from the visited set
                // so that the same object can be accessed in different branches (A -> B, A -> C).
                if (!type.IsValueType)
                    visited.Remove(value);
            }
        }
    }

    /// <summary>
    /// Hooks into the sheet processing pipeline to handle dynamic sheet generation based on template placeholders.
    /// If a sheet name matches a specific pattern and the corresponding input value is an enumerable,
    /// it generates multiple sheets based on the elements of the enumerable and returns true.
    /// </summary>
    [CreateSyncVersion]
    private async Task<bool> TryExpandParametrizedSheetAsync(OpenXmlZip outputFileArchive, string originalSheetName, IDictionary<int, string> templateSharedStrings, int sheetIndex, List<(int Index, string Name)> allSheetInfos, ZipArchiveEntry templateSheet, IDictionary<string, object?> inputValues, CancellationToken cancellationToken = default)
    {
        // Use regex to match the sheet name to pattern "$PlaceholderName$"
        var match = ParametrizedSheetRegexImpl.Match(originalSheetName);

        // Check if the pattern matches, the placeholder exists in input values, and the value is an IEnumerable
        if (!match.Success ||
            !inputValues.TryGetValue(match.Groups[1].Value, out var subObj) || 
            subObj is not IEnumerable subIter)
        {
            return false;
        }

        // Extract the base sheet name from the template placeholder
        var baseSheetName = match.Groups[1].Value;
        var subIndex = 1;

        // 1. Batch create all worksheet files
        foreach (var subRoot in subIter)
        {
            // Clear internal state collections before processing each new sheet
            _xRowInfos.Clear();
            _xMergeCellInfos.Clear();
            _newXMergeCellInfos.Clear();
            _calcChainCellRefs.Clear();

            // Extract values for the current iteration item into a dictionary
            var subValues = _inputValueExtractor.ToValueDictionary(subRoot);

            // Define the internal path for the new sheet XML file
            var newSheetPath = $"xl/worksheets/sheet{sheetIndex}.xml";

            // Check if a custom "SheetName" was provided in the current item's values, or fallback base name + index
            var finalSheetName = subValues.TryGetValue("SheetName", out var customSheetName) && customSheetName is not null
                ? customSheetName.ToString()?.Trim() ?? $"{baseSheetName}{subIndex++}"
                : $"{baseSheetName}{sheetIndex}";

            // Only collect sheet info, do not call configuration methods yet
            allSheetInfos.Add((sheetIndex, finalSheetName));

            // Create the new worksheet entry in the output ZIP archive
            var newSheetEntry = outputFileArchive.ZipFile.CreateEntry(newSheetPath);
            var newSheetStream = await newSheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableSheetStream = newSheetStream.ConfigureAwait(false);

            // Generate the sheet content based on the template and current sub-values
            await GenerateSheetByCreateModeAsync(templateSheet, newSheetStream, subValues, templateSharedStrings, cancellationToken: cancellationToken).ConfigureAwait(false);

            // Append calculation chain content for the newly created sheet
            _calcChainContent.Append(CalcChainHelper.GetCalcChainContent(_calcChainCellRefs, sheetIndex));
        }

        return true;
    }

    /// <summary>
    /// Adds worksheets to the workbook and register them int workbook.xml and workbook.xml.rels
    /// </summary>
    [CreateSyncVersion]
    private static async Task BatchAddSheetsToWorkbookAsync(ZipArchive outputZip, ZipArchive templateArchive, List<(int Index, string Name)> sheetInfos, CancellationToken cancellationToken)
    {
        // Load the workbook and its relationships from the template
        var relDoc = await LoadXmlAsync(templateArchive, ExcelFileNames.WorkbookRels, cancellationToken).ConfigureAwait(false);
        var wbDoc = await LoadXmlAsync(templateArchive, ExcelFileNames.Workbook, cancellationToken).ConfigureAwait(false);

        // 1. Clear all existing <sheet> elements in workbook.xml to rebuild a clean <sheets> container
        if (wbDoc.Root?.Element(SpreadsheetNs + "sheets") is { } sheetsPart)
        {
            // Directly remove child nodes, keeping the container and default namespaces 
            sheetsPart.Elements().Remove();
        }
        else
        {
            // If the original template lacks a sheets node, create a new one and append it to the root
            wbDoc.Root?.Add(new XElement(SpreadsheetNs + "sheets"));
        }

        // 2.  Clean up all relationship records pointing to worksheets in workbook.xml.rels
        var relsRoot = relDoc.Root;
        if (relsRoot != null)
        {
            // Only delete relationships of Type 'worksheet', preserving core relationships like sharedStrings/styles/theme
            var worksheetRels = relsRoot.Elements(PackageRelNs + "Relationship")
                .Where(r => r.Attribute("Type")?.Value == Schemas.SpreadsheetmlXmlWorksheetRelationship);

            // Remove the filtered worksheet relationships
            foreach (var rel in worksheetRels) 
                rel.Remove();
        }

        // Batch add new relationship records for each generated sheet
        foreach (var sheet in sheetInfos)
        {
            relDoc.Root!.Add(new XElement(PackageRelNs + "Relationship",
                new XAttribute("Id", $"rIdSheet{sheet.Index}"),
                new XAttribute("Type", Schemas.SpreadsheetmlXmlWorksheetRelationship),
                new XAttribute("Target", $"worksheets/sheet{sheet.Index}.xml")));
        }

        // Batch add new sheet definitions to the workbook XML
        var sheetsNode = wbDoc.Descendants(SpreadsheetNs + "sheets").FirstOrDefault();
        if (sheetsNode != null)
        {
            foreach (var sheet in sheetInfos)
            {
                sheetsNode.Add(new XElement(SpreadsheetNs + "sheet",
                    new XAttribute("name", sheet.Name),
                    new XAttribute("sheetId", sheet.Index),
                    new XAttribute(SpreadsheetRelNs + "id", $"rIdSheet{sheet.Index}")));
            }
        }

        // Save the modified xml entries
        await SaveXmlToZipAsync(outputZip, ExcelFileNames.WorkbookRels, relDoc, cancellationToken).ConfigureAwait(false);
        await SaveXmlToZipAsync(outputZip, ExcelFileNames.Workbook, wbDoc, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Parses the template to build a mapping from each sheet name to the corresponding xml path
    /// </summary>
    [CreateSyncVersion]
    private async Task<Dictionary<string, string>> GetSheetNameMapAsync(ZipArchive archive, CancellationToken cancellationToken = default)
    {
        Dictionary<string, string> nameToPath = [];
        Dictionary<string, string> ridToSheetPath = [];

        // 1. Read workbook.xml.rels to get the mapping of rId -> sheet path
        if (await LoadXmlAsync(archive, ExcelFileNames.WorkbookRels, cancellationToken).ConfigureAwait(false) is not { } relDoc)
            return [];

        foreach (var rel in relDoc.Descendants(PackageRelNs + "Relationship"))
        {
            if (rel.Attribute("Id")?.Value is { } rid)
            {
                var target = rel.Attribute("Target")?.Value;
                if (string.IsNullOrEmpty(rid) || string.IsNullOrEmpty(target)) continue;

                // Construct the full internal path (ensure forward slashes for consistency)
                var fullSheetPath = Path.Combine("xl", target).Replace("\\", "/");
                ridToSheetPath[rid] = fullSheetPath;
            }
        }

        // 2. Read workbook.xml to get the Real Sheet Name + rId mapping
        if (await LoadXmlAsync(archive, ExcelFileNames.Workbook, cancellationToken).ConfigureAwait(false) is not { } wbDoc)
            return [];

        foreach (var sheetNode in wbDoc.Descendants(SpreadsheetNs + "sheet"))
        {
            var realName = sheetNode.Attribute("name")?.Value.Trim();
            var rid = sheetNode.Attribute(SpreadsheetRelNs + "id")?.Value;
            if (string.IsNullOrEmpty(realName) || string.IsNullOrEmpty(rid))
                continue;

            // If the rId exists in our temporary mapping, link the XML path to the real name
            if (ridToSheetPath.TryGetValue(rid!, out var sheetPath))
            {
                // key: xml path, value: real sheet name
                nameToPath[sheetPath] = realName!;
            }
        }

        return nameToPath;
    }

    [CreateSyncVersion]
    private static async Task<XDocument> LoadXmlAsync(ZipArchive templateArchive, string path, CancellationToken cancellationToken)
    {
        var entry = templateArchive.GetEntry(path)!;
        var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);

        return await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task SaveXmlToZipAsync(ZipArchive outputZip, string path, XDocument doc, CancellationToken cancellationToken)
    {
        var newEntry = outputZip.CreateEntry(path);
        var stream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);

        await doc.SaveAsync(stream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
    }
}
