using System.Buffers;
using System.Collections.Concurrent;
using ClosedXML.Excel;

namespace MiniExcelLib.Tests.TemplateOptimization;

#region Models

/// <summary>
/// Represents access types for file system permissions
/// </summary>
public enum AccessType : byte
{
    None = 0,
    FullControl = 1, // 'F'
    Modify = 2, // 'M'
    Write = 3, // 'W'
    Read = 4, // 'R'
    Special = 5 // 'S'
}

public class AccountInfo
{
    public SecurityIdentifier Sid { get; set; }
    public string Name { get; set; }
    public string Domain { get; set; }
}

#endregion

#region Helper

/// <summary>
/// Helper class for account-related operations
/// </summary>
public static class ReportHelper
{
    public const double MaxMatrixIdentityPart = 400;
    public const double MaxMatrixRow = 400;

    public readonly struct MatrixRow(string path, AccessType[] access)
    {
        public string Path { get; } = path;
        public AccessType[] Access { get; } = access;

        public Dictionary<string, object> ToDictionary(int startPart, int endPart)
        {
            var dict = new Dictionary<string, object?>(endPart - startPart + 1)
            {
                ["Path"] = Path
            };

            for (var i = startPart; i < endPart; i++)
            {
                if (i < Access.Length)
                {
                    dict[i.ToString()] = GetAccessType(Access[i]);
                }
            }

            return dict;
        }
    }

    private static ConcurrentDictionary<FileSystemRights, AccessType> AccessTypeMap { get; } = new(
        new Dictionary<FileSystemRights, AccessType>
        {
            { FileSystemRights.FullControl, AccessType.FullControl },
            { FileSystemRights.Modify | FileSystemRights.Synchronize, AccessType.Modify },
            { FileSystemRights.Write | FileSystemRights.Synchronize, AccessType.Write },
            { FileSystemRights.Read | FileSystemRights.ReadAndExecute | FileSystemRights.Synchronize, AccessType.Read }
        });

    public static string? GetAccessType(AccessType accessType) => accessType switch
    {
        AccessType.None => null,
        AccessType.FullControl => "F",
        AccessType.Modify => "M",
        AccessType.Write => "W",
        AccessType.Read => "R",
        AccessType.Special => "S",
        _ => string.Empty
    };

    public static AccessType GetAccessType(FileSystemRights acl) => AccessTypeMap.GetValueOrDefault(acl, AccessType.Special);

    public static IEnumerable<T> Page<T>(IEnumerable<T> en, int pageSize, int page)
    {
        return en.Skip(page * pageSize).Take(pageSize);
    }
}

#endregion

public class MatrixCreator
{
    private ConcurrentDictionary<string, AccessType[]> _entryCache;
    private readonly List<FileSystemEntry> _resultList;
    private readonly List<SecurityIdentifier> _aclSids;
    private readonly ConcurrentDictionary<SecurityIdentifier, int> _identityCache;

    public MatrixCreator(List<FileSystemEntry> resultList, List<SecurityIdentifier> aclSids)
    {
        _entryCache = [];
        _resultList = resultList;
        _aclSids = aclSids;
        _identityCache = new (_aclSids
            .Distinct()
            .Select((sid, index) => new { sid, index })
            .ToDictionary(x => x.sid, x => x.index)
        );
    }


    public IEnumerable<ReportHelper.MatrixRow> Creation()
    {
        var first = _resultList.First();
        yield return new ReportHelper.MatrixRow(first.Path, CheckPathAccess(first.Path, first.Acl));

        foreach (var fileSystemStruct in _resultList)
        {
            var child = fileSystemStruct.Path;

            var childAccessArray = CheckPathAccess(child, fileSystemStruct.Acl);
            var parentArray = CheckPathAccess(Path.GetDirectoryName(child) ?? child);


            // If the parent and current Path array is not sequence equal
            if (AreAccessArraysEqual(childAccessArray, parentArray)) continue;
            yield return new ReportHelper.MatrixRow(child, childAccessArray);
        }
    }

    private static bool AreAccessArraysEqual(AccessType[] array1, AccessType[] array2)
    {
        // Quick check: if both arrays have the same reference, they're equal
        return MemoryExtensions.SequenceEqual<AccessType>(array1, array2);
    }

    private AccessType[] CheckPathAccess(string currentFolder,
        Dictionary<SecurityIdentifier, FileSystemRights>? accessDict = null)
    {
        if (_entryCache.TryGetValue(currentFolder, out var accessArray)) return accessArray;
        if (accessDict is null) return [];
        var accessSpan = GetAccess(accessDict);
        if (!accessSpan.TryCopyTo(accessArray) || accessArray is null) accessArray = accessSpan.ToArray();
        _entryCache.TryAdd(currentFolder, accessArray);

        return accessArray;
    }

    private Span<AccessType> GetAccess(Dictionary<SecurityIdentifier, FileSystemRights> accessDict)
    {
        var pooledArray = ArrayPool<AccessType>.Shared.Rent(_identityCache.Count);
        try
        {
            var resultSpan = pooledArray.AsSpan(0, _identityCache.Count);
            resultSpan.Clear();

            foreach (var (sid, rights) in accessDict)
            {
                if (!_identityCache.TryGetValue(sid, out var index)) continue;
                resultSpan[index] = ReportHelper.GetAccessType(rights);
            }

            return resultSpan;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in matrix row creation: {ex.Message}");
            return [];
        }
        finally
        {
            // Return the array to the pool
            ArrayPool<AccessType>.Shared.Return(pooledArray);
        }
    }

    public string ExcelFileCreator(string outputPath)
    {
        Console.WriteLine("Creating base template");
        // Step 1: Create template with ClosedXML
        var templatePath = Path.Combine(Path.GetDirectoryName(outputPath) ?? string.Empty, "template.xlsx");
        using var workbook = new XLWorkbook();

        MatrixSheet(workbook);

        workbook.SaveAs(templatePath);
        
        return templatePath;
    }


    private void MatrixSheet(XLWorkbook workbook)
    {
        var matrixRowCount = Creation().Count();
        var pageCount = (int)Math.Ceiling(matrixRowCount / ReportHelper.MaxMatrixRow);
        if (pageCount == 0 || matrixRowCount == 1)
        {
            Console.WriteLine("No data found. Matrix will be skipped.");
            return;
        }

        var identityParts = (int)Math.Ceiling(_identityCache.Count / ReportHelper.MaxMatrixIdentityPart);
        if (_identityCache.Count > ReportHelper.MaxMatrixIdentityPart)
            Console.WriteLine($"Warning: Currently found users count is {_identityCache.Count}. " +
                              $"Will be separated into {identityParts} parts per page ({pageCount}/{identityParts}).");

        for (var page = 0; page < pageCount; page++)
        {
            for (var part = 0; part < identityParts; part++)
            {
                var partName = identityParts == 1 ? "" : $"-{part + 1}";
                var sheetName = "Matrix" + (page == 0 && identityParts == 1 ? "" : $"_{page + 1}{partName}");

                var matrixSheet = workbook.Worksheets.Add(sheetName);
                Console.WriteLine($"Creating {sheetName} sheet");
                // Create header row
                var userIndex = 2;

                //Set up Rows
                var headerRow = matrixSheet.Row(1);
                var fillDataRow = matrixSheet.Row(2);
                fillDataRow.Cell(1).Value = $"{{{{{sheetName}.Path}}}}";

                // Set up splitter cell first
                var splitterCell = matrixSheet.Cell(1, 1);
                splitterCell.Value = "Path";
                splitterCell.Style.Font.SetFontColor(XLColor.ForestGreen);
                matrixSheet.Column(1).Width = 60;

                var headerStyle = headerRow.Cell(userIndex).Style;
                headerStyle.Font.SetFontColor(XLColor.White);
                headerStyle.Font.SetBold(true);
                headerStyle.Alignment.SetTextRotation(90);
                headerStyle.Fill.SetBackgroundColor(XLColor.ForestGreen);

                foreach (var sid in ReportHelper.Page(_identityCache, (int)ReportHelper.MaxMatrixIdentityPart, part))
                {
                    var header = headerRow.Cell(userIndex);


                    header.Style = headerStyle;
                    header.Value = $"{sid.Key.Value}";
                    fillDataRow.Cell(userIndex).Value =
                        $"{{{{{sheetName}.{part * ReportHelper.MaxMatrixIdentityPart + userIndex - 2}}}}}";
                    matrixSheet.Column(userIndex).Width = 3;
                    userIndex++;
                }

                headerRow.SetAutoFilter();
                matrixSheet.SheetView.Freeze(1, 1);
            }
        }
    }
}