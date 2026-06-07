using System.ComponentModel;
using System.Xml.Linq;

namespace MiniExcelLib.OpenXml.Templates;

/// <summary>
/// To Support 
/// 
/// public record Identity(int Type, string No);
/// 
/// public class Person
/// {
///     public Identity Id { get; set; }
///     public string Name { get; set; }
/// }
/// 
/// var obj = new { p = new Person { Id = new Identity(1, "A123"), Name = "张三" }, ps = [..some Person] };
/// 
/// in the template, you can write:
/// case 1 
///     {{p.Id.Type}}}, {{p.Id.No}}}, {{p.Name}}
/// 
/// case 2
/// set sheet name to $ps$ and template
/// {{Id.Type}}}, {{Id.No}}}, {{Name}} 
/// then it will generate multiple sheets based on the elements in ps, and each sheet will have the corresponding values of Id.Type, Id.No, Name for that element.
/// </summary>
internal partial class OpenXmlTemplate
{

#if !NET8_0_OR_GREATER
    /// <summary>
    /// Custom equality comparer that uses reference equality instead of overridden object.Equals.
    /// Required for .NET versions prior to 8.0 where ReferenceEqualityComparer is not built-in.
    /// </summary>
    public class ReferenceComparer : IEqualityComparer<object>
    {
        /// <summary>
        /// Determines whether the specified objects are the exact same instance in memory.
        /// </summary>
        public new bool Equals(object x, object y) => ReferenceEquals(x, y);

        /// <summary>
        /// Returns a hash code based on the object's memory reference.
        /// </summary>
        public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
    }
#endif

    /// <summary>Default max recursion depth (can be adjusted based on business needs)</summary>
    private const int DefaultMaxDepth = 4;

    /// <summary>
    /// Recursively flattens an object graph into a dictionary of "key.subkey" pairs and fully formats the values.
    /// Includes protection against circular references and stack overflow via depth limiting.
    /// </summary>
    /// <param name="replacements">The target dictionary to store the flattened key-value results.</param>
    /// <param name="key">The current key prefix for the nested property.</param>
    /// <param name="value">The current object value to process.</param>
    /// <param name="propInfo">The PropertyInfo of the current property, used for reading custom formatting attributes.</param>
    /// <param name="maxDepth">The maximum allowed recursion depth. Falls back to ToString() when exceeded.</param>
    public static void AddFlattenedAndFormattedValues(
        Dictionary<string, string> replacements,
        string key,
        object? value,
        PropertyInfo? propInfo = null,
        int maxDepth = DefaultMaxDepth)
    {
        // Initialize a HashSet to track visited objects and prevent infinite loops from circular references.
        // Use reference equality comparer to prevent misjudgments caused by business types overriding Equals/GetHashCode.
#if NET8_0_OR_GREATER
        var visited = new HashSet<object>(ReferenceEqualityComparer.Instance); 
#else
        var visited = new HashSet<object>(new ReferenceComparer());
#endif

        // Start the recursive core processing with initial depth set to 0.
        Core(replacements, key, value, propInfo, maxDepth, 0, visited);
    }

    /// <summary>
    /// The internal recursive method that performs the actual object traversal, flattening, and formatting.
    /// </summary>
    private static void Core(
        Dictionary<string, string> replacements,
        string key,
        object? value,
        PropertyInfo? propInfo,
        int maxDepth,
        int currentDepth,
        HashSet<object> visited)
    {
        // Handle null values or invalid types by assigning an empty string to the current key and exiting early.
        if (value == null || value.GetType() is not Type type)
        {
            replacements[key] = string.Empty;
            return;
        }

        // 1. Primitive types / Enums: Format directly, do not consume depth, and do not enter reference tracking.
        if (IsSimpleType(type) || type.IsEnum)
        {
            replacements[key] = GetFormattedValue(propInfo, value, type);
            return;
        }

        // 2. Depth control: Safe fallback to string representation when exceeding the limit to avoid OOM/StackOverflow.
        if (currentDepth >= maxDepth)
            return;

        // 3. Circular reference detection (only for reference types; value types cannot form reference loops).
        // If the object is already in the visited set, it's a circular reference.
        if (!type.IsValueType && !visited.Add(value))
            return;

        try
        {
            // 4. Dictionary handling: Iterate through key-value pairs and recursively process values.
            if (value is Dictionary<string, object> dict)
            {
                foreach (var kv in dict)
                {
                    // Construct the new sub-key by appending the dictionary key.
                    var subKey = string.Concat(key, ".", kv.Key);
                    // Recursively call Core for the dictionary value.
                    Core(replacements, subKey, kv.Value, propInfo, maxDepth, currentDepth + 1, visited);
                }
                return;
            }

            // add self
            replacements[key] = GetFormattedValue(propInfo, value, type);

            // 5. Object property recursion: Reflect over public instance properties.
            // Filter out indexers and write-only properties.

            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
               .Where(p => p.CanRead && p.GetIndexParameters().Length == 0);

            foreach (var prop in properties)
            {
                // Construct the new sub-key by appending the property name.
                var subKey = string.Concat(key, ".", prop.Name);
                // Retrieve the actual value of the property from the current object instance.
                var subValue = prop.GetValue(value);
                // Recursively call Core for the property value.
                Core(replacements, subKey, subValue, prop, maxDepth, currentDepth + 1, visited);
            }
        }
        finally
        {
            // Remove the current node from the visited set.
            // This allows the same object to be accessed normally in different branches (DAG shared references).
            // It only intercepts true "loops" (A -> B -> A) and does not kill legitimate object reuse (A -> B, A -> C).
            if (!type.IsValueType)
                visited.Remove(value);
        }
    }

    #region Formatting Logic (Maintained independently, core behavior unchanged)

    /// <summary>
    /// Formats the given cell value into a string representation suitable for OpenXml injection.
    /// Handles specific types like booleans, dates, enums, and numeric values.
    /// </summary>
    private static string GetFormattedValue(PropertyInfo? propInfo, object? cellValue, Type type)
    {
        // Variable to hold the intermediate string representation of the cell value.
        string? cellValueStr;

        // handle as original write
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
            // Get the string name of the enum value.
            var stringValue = Enum.GetName(type, cellValue!) ?? "";
            // Retrieve the DescriptionAttribute from the enum field.
            var attr = type.GetField(stringValue)?.GetCustomAttribute<DescriptionAttribute>();
            // Use the description if it exists, otherwise fallback to the enum string name.
            var description = attr?.Description ?? stringValue;
            // Encode the final string to ensure it is safe for XML.
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

    /// <summary>
    /// Determines if a given type is a simple/primitive type that should not be recursively traversed.
    /// </summary>
    private static bool IsSimpleType(Type type)
    {
        return type.IsPrimitive || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime) || type == typeof(Guid) || Nullable.GetUnderlyingType(type) != null;
    }

    #endregion

    /// <summary>
    /// Hooks into the sheet processing pipeline to handle dynamic sheet generation based on template placeholders.
    /// If a sheet name matches a specific pattern and the corresponding input value is an enumerable,
    /// it generates multiple sheets based on the elements of the enumerable.
    /// </summary>
    /// <param name="outputFileArchive">The output ZIP archive where new sheets will be created.</param>
    /// <param name="realSheetName">The original name of the sheet from the template.</param>
    /// <param name="templateSharedStrings">Dictionary of shared strings from the template.</param>
    /// <param name="sheetIdx">The current sheet index counter.</param>
    /// <param name="allSheetInfos">List to collect information about all generated sheets for later workbook configuration.</param>
    /// <param name="templateSheet">The ZIP entry of the template sheet to clone.</param>
    /// <param name="templateFullName">The full name/path of the template sheet.</param>
    /// <param name="inputValues">The dictionary of input values provided for template rendering.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>True if dynamic sheets were generated, otherwise false.</returns>
    private async Task<bool> HookSheetProcess(OpenXmlZip outputFileArchive, string realSheetName, IDictionary<int, string> templateSharedStrings, int sheetIdx, List<(int Index, string Name)> allSheetInfos, ZipArchiveEntry templateSheet, string templateFullName, IDictionary<string, object?> inputValues, CancellationToken cancellationToken)
    {
        // Use regex to match the sheet name pattern "$PlaceholderName$"
        var m = Regex.Match(realSheetName, @"\$([^$]+)\$");

        // Check if the pattern matches, the placeholder exists in input values, and the value is an IEnumerable
        if (m.Success && inputValues.TryGetValue(m.Groups[1].Value, out var subObj) && subObj is IEnumerable sunIter)
        {
            // Extract the base sheet name from the template placeholder
            var baseSheetName = m.Groups[1].Value;
            var subIndex = 1; // Sub-index for naming sheets if custom names are not provided

            // 1. Batch create all worksheet files first (streams are automatically closed, preventing conflicts)
            foreach (var subRoot in sunIter)
            {
                // Clear internal state collections before processing each new sheet
                _xRowInfos.Clear();
                _xMergeCellInfos.Clear();
                _newXMergeCellInfos.Clear();
                _calcChainCellRefs.Clear();

                // Extract values for the current iteration item into a dictionary
                var subValues = _inputValueExtractor.ToValueDictionary(subRoot);

                // Increment the global sheet index
                sheetIdx++;
                // Define the internal path for the new sheet XML file
                var newSheetPath = $"xl/worksheets/sheet{sheetIdx}.xml";

                // Determine the final sheet name
                string finalSheetName;
                // Check if a custom "SheetName" is provided in the current item's values
                if (subValues.TryGetValue("SheetName", out var customSheetName) && customSheetName != null)
                {
                    finalSheetName = customSheetName.ToString()?.Trim() ?? $"{baseSheetName}{subIndex++}";
                }
                else
                {
                    // Fallback to base name + index if no custom name is provided
                    finalSheetName = $"{baseSheetName}{sheetIdx++}";
                }

                // Only collect sheet info, do not call configuration methods yet
                allSheetInfos.Add((sheetIdx, finalSheetName));

                // Create the new worksheet entry in the output ZIP archive
                var newSheetEntry = outputFileArchive.ZipFile.CreateEntry(newSheetPath);

                // Open the stream for the new sheet entry (handling .NET 8+ async differences)
#if NET8_0_OR_GREATER
                using var newSheetStream = await newSheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
                using var newSheetStream = newSheetEntry.Open();
#endif
                // Generate the sheet content based on the template and current sub-values
                await GenerateSheetByCreateModeAsync(templateSheet, newSheetStream, subValues, templateSharedStrings, cancellationToken: cancellationToken).ConfigureAwait(false);

                // Append calculation chain content for the newly created sheet
                _calcChainContent.Append(CalcChainHelper.GetCalcChainContent(_calcChainCellRefs, sheetIdx));
            }

            // Return true to indicate that dynamic sheets were successfully processed
            return true;
        }

        // Return false if the sheet name did not match the dynamic generation pattern
        return false;
    }

    /// <summary>
    /// Batch adds worksheets to the Excel configuration (writes all at once, no overwriting, ensures sheet names take effect).
    /// Modifies workbook.xml and workbook.xml.rels to register the newly created sheets.
    /// </summary>
    /// <param name="outputZip">The output ZIP archive being generated.</param>
    /// <param name="templateArchive">The original template ZIP archive.</param>
    /// <param name="sheetInfos">List of tuples containing the sheet index and final name.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    [CreateSyncVersion]
    public static async Task BatchAddSheetsToExcelConfigAsync(ZipArchive outputZip, ZipArchive templateArchive, List<(int Index, string Name)> sheetInfos, CancellationToken cancellationToken)
    {
        // ======================================
        // Phase 1: Pure in-memory reading and modification (all streams are closed immediately after reading)
        // ======================================

        // Load the relationships XML from the template
        XDocument relDoc = await LoadTemplateXmlAsync(templateArchive, "xl/_rels/workbook.xml.rels", cancellationToken).ConfigureAwait(false);
        // Load the main workbook XML from the template
        XDocument wbDoc = await LoadTemplateXmlAsync(templateArchive, "xl/workbook.xml", cancellationToken).ConfigureAwait(false);

        // Define standard OpenXml namespaces
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // 1. Clear all existing <sheet> elements in workbook.xml to rebuild a clean <sheets> container
        var sheetsPart = wbDoc.Root?.Element(ssNs + "sheets");
        if (sheetsPart != null)
        {
            // Directly remove child nodes, keeping the container itself and default namespaces 
            // (to avoid Excel errors caused by missing namespaces)
            sheetsPart.Elements().Remove();
        }
        else
        {
            // If the original template lacks a sheets node, create a new one and append it to the root
            wbDoc.Root?.Add(new XElement(ssNs + "sheets"));
        }

        // 2.  Clean up all relationship records pointing to worksheets in workbook.xml.rels
        const string worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

        var relsRoot = relDoc.Root;
        if (relsRoot != null)
        {
            // Only delete relationships of Type 'worksheet', preserving core relationships like sharedStrings/styles/theme
            var worksheetRels = relsRoot.Elements(relNs + "Relationship")
                .Where(r => r.Attribute("Type")?.Value == worksheetRelType)
                .ToList();

            // Remove the filtered worksheet relationships
            foreach (var rel in worksheetRels) rel.Remove();
        }

        // Batch add new relationship records for each generated sheet
        foreach (var sheet in sheetInfos)
        {
            relDoc.Root!.Add(new XElement(relNs + "Relationship",
                new XAttribute("Id", $"rIdSheet{sheet.Index}"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                new XAttribute("Target", $"worksheets/sheet{sheet.Index}.xml")));
        }

        // Batch add new sheet definitions to the workbook XML
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
        // Phase 2: All streams are closed → Safely create entries and write modified XML back to the ZIP
        // ======================================

        // Save the modified relationships XML
        await SaveXmlToZipAsync(outputZip, "xl/_rels/workbook.xml.rels", relDoc, cancellationToken).ConfigureAwait(false);
        // Save the modified workbook XML
        await SaveXmlToZipAsync(outputZip, "xl/workbook.xml", wbDoc, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads an XML file from the template archive into memory (stream is automatically closed, returns an in-memory XDocument).
    /// </summary>
    /// <param name="templateArchive">The template ZIP archive.</param>
    /// <param name="path">The internal path of the XML file to load.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    /// <returns>The loaded XDocument.</returns>
    [CreateSyncVersion]
    private static async Task<XDocument> LoadTemplateXmlAsync(ZipArchive templateArchive, string path, CancellationToken cancellationToken)
    {
        // Retrieve the ZIP entry for the specified path
        var entry = templateArchive.GetEntry(path)!;

        // Open the stream and load the XDocument, handling .NET 8+ async differences
#if NET8_0_OR_GREATER
        using var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        return await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var stream = entry.Open();
        return XDocument.Load(stream);
#endif
    }

    /// <summary>
    /// Writes an XDocument to a specified path within the output ZIP archive (stream is automatically closed, no residual locks).
    /// </summary>
    /// <param name="outputZip">The output ZIP archive.</param>
    /// <param name="path">The internal path where the XML file should be saved.</param>
    /// <param name="doc">The XDocument to save.</param>
    /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
    [CreateSyncVersion]
    private static async Task SaveXmlToZipAsync(ZipArchive outputZip, string path, XDocument doc, CancellationToken cancellationToken)
    {
        // At this point, no streams are open, so it is safe to create a new entry

        // Create the new ZIP entry
        var newEntry = outputZip.CreateEntry(path);

        // Open the stream and save the XDocument, handling .NET 8+ async differences
#if NET8_0_OR_GREATER
        using var stream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await doc.SaveAsync(stream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var stream = newEntry.Open();
        doc.Save(stream);
#endif
    }

    /// <summary>
    /// Accurately parses the template to build a mapping of [Real Sheet Name] to [Corresponding Sheet XML Path].
    /// </summary>
    /// <param name="archive">The ZIP archive to parse.</param>
    /// <returns>A dictionary where the key is the sheet XML path and the value is the real sheet name.</returns>
    public Dictionary<string, string> GetRealSheetNameMap(ZipArchive archive)
    {
        // Dictionary to hold the final mapping of XML path -> Real Sheet Name
        var nameToPath = new Dictionary<string, string>();
        // Temporary dictionary to hold the mapping of Relationship Id -> Sheet XML Path
        var ridToSheetPath = new Dictionary<string, string>();

        // 1. Read workbook.xml.rels to get the mapping of rId -> sheet file path
        var relsEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/_rels/workbook.xml.rels");
        if (relsEntry == null) return nameToPath; // Return empty if relationships file is missing

        using var relStream = relsEntry.Open();
        var relDoc = XDocument.Load(relStream);
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";

        // Iterate through all Relationship elements
        foreach (var rel in relDoc.Descendants(relNs + "Relationship"))
        {
            var rid = rel.Attribute("Id")?.Value;
            var target = rel.Attribute("Target")?.Value;
            if (string.IsNullOrEmpty(rid) || string.IsNullOrEmpty(target)) continue;

            // Construct the full internal path (ensure forward slashes for consistency)
            var fullSheetPath = Path.Combine("xl", target).Replace("\\", "/");
            ridToSheetPath[rid] = fullSheetPath;
        }

        // 2. Read workbook.xml to get the Real Sheet Name + rId mapping
        var wbEntry = archive.Entries.FirstOrDefault(e => e.FullName == "xl/workbook.xml");
        if (wbEntry == null) return nameToPath; // Return empty if workbook file is missing

        using var wbStream = wbEntry.Open();
        var wbDoc = XDocument.Load(wbStream);
        XNamespace ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // Iterate through all sheet definitions in the workbook
        foreach (var sheetNode in wbDoc.Descendants(ssNs + "sheet"))
        {
            var realName = sheetNode.Attribute("name")?.Value?.Trim() ?? "";
            var rid = sheetNode.Attribute(rNs + "id")?.Value ?? "";
            if (string.IsNullOrEmpty(realName) || string.IsNullOrEmpty(rid)) continue;

            // If the rId exists in our temporary mapping, link the XML path to the real name
            if (ridToSheetPath.TryGetValue(rid, out var sheetPath))
            {
                nameToPath[sheetPath] = realName; // key: xml path, value: real sheet name
            }
        }

        // Return the completed mapping
        return nameToPath;
    }
}