using ClosedXML.Excel;

namespace MiniExcelLib.Tests.TemplateOptimization;

internal static class MiniExcelVersion
{
    private static IEnumerable<T> Page<T>(IEnumerable<T> en, int pageSize, int page)
    {
        return en.Skip(page * pageSize).Take(pageSize);
    }

    private static IEnumerable<object> OptimizeFileSystemStructs(IEnumerable<FileSystemEntry> fileSystemStructs)
    {
        return fileSystemStructs.SelectMany(entry =>
            entry.Acl.Count != 0
                ? entry.Acl.Select(acl => CreateRawEntry(entry, acl.Key.ToString(), acl.Value.ToString()))
                : [CreateRawEntry(entry, string.Empty, string.Empty)]
        );
    }

    private static object CreateRawEntry(FileSystemEntry entry, string aclIdentity, string aclRights)
    {
        return new
        {
            entry.Path,
            entry.Owner,
            entry.FileEntryAttributes,
            ACLIdentity = aclIdentity,
            ACLRights = aclRights,
            entry.IsModified,
            entry.Error
        };
    }

    public static void RawDataFill(IEnumerable<FileSystemEntry> fileSystemStructs, string filePath, int maxRowPerSheet)
    {
        var result = OptimizeFileSystemStructs(fileSystemStructs);

        var sheets = new Dictionary<string, object>();
        var pageNumbers = Math.Ceiling((decimal)(result.Count() / (double)maxRowPerSheet));
        for (var pager = 0; pager < pageNumbers; pager++)
        {
            sheets.TryAdd("RawData " + (pager == 0 ? "" : pager + 1), Page(result, maxRowPerSheet, pager));
        }

        using var stream = File.Create(filePath);
        MiniExcel.Exporter.GetOpenXmlExporter().Export(stream, sheets, configuration: new OpenXmlConfiguration
        {
            FreezeRowCount = 1, DynamicColumns =
            [
                new DynamicExcelColumn("Path") { Width = 65 },
                new DynamicExcelColumn("Owner") { Width = 50 },
                new DynamicExcelColumn("FileEntryAttributes") { Width = 20 },
                new DynamicExcelColumn("ACLIdentity") { Name = "Account Identity", Width = 50 },
                new DynamicExcelColumn("ACLRights") { Name = "Account Rights", Width = 30 },
                new DynamicExcelColumn("IsModified") { Name = "Modified" },
                new DynamicExcelColumn("Error")
            ]
        });
        sheets.Clear();

        using var workbook = new XLWorkbook(stream);
        foreach (var worksheet in workbook.Worksheets)
        {
            var firstRow = worksheet.FirstRow();
            var firstRowStyle = firstRow.Style;
            firstRowStyle.Fill.SetBackgroundColor(XLColor.FromHtml("#FF228B22"));
            firstRowStyle.Font.Bold = true;
            firstRow.Height = 30;
            firstRowStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            firstRowStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            firstRowStyle.Font.SetFontColor(XLColor.White);
        }
    }
}

internal static class ClosedXMLVersion
{
    private static IEnumerable<T> Page<T>(IEnumerable<T> en, int pageSize, int page)
    {
        return en.Skip(page * pageSize).Take(pageSize);
    }

    private static IEnumerable<object> OptimizeFileSystemStructs(IEnumerable<FileSystemEntry> fileSystemStructs)
    {
        return fileSystemStructs.SelectMany(entry =>
            entry.Acl.Count != 0
                ? entry.Acl.Select(acl => CreateRawEntry(entry, acl.Key.ToString(), acl.Value.ToString()))
                : [CreateRawEntry(entry, string.Empty, string.Empty)]
        );
    }

    private static object CreateRawEntry(FileSystemEntry entry, string aclIdentity, string aclRights) => new
    {
        entry.Path,
        entry.Owner,
        entry.FileEntryAttributes,
        ACLIdentity = aclIdentity,
        ACLRights = aclRights,
        entry.IsModified,
        entry.Error
    };

    private enum HeaderColumn
    {
        Folder,
        Owner,
        FileEntryAttributes,
        ACLUser,
        ACLAuthorityLevel,
        Modified,
        Error
    }


    /// <summary>
    /// This method populates the 'rawdata' sheet in the Excel workbook with the provided results list.
    /// </summary>
    /// <param name="excel">The Excel workbook to work with.</param>
    /// <param name="resultsList">The list of results to populate the sheet with.</param>
    /// <summary>
    public static void RawDataFill(IEnumerable<FileSystemEntry> fileSystemEntries, string filePath, int maxRowPerSheet)
    {
        using var workbook = !File.Exists(filePath) ? new XLWorkbook() : new XLWorkbook(filePath);
        var result = OptimizeFileSystemStructs(fileSystemEntries);
        var maxPages = Math.Ceiling((decimal)(result.Count() / (double)maxRowPerSheet));

        for (var pager = 0; pager < maxPages; pager++)
        {
            var worksheetName = "RawData" + (pager == 0 ? "" : $" {pager + 1}");
            var worksheet = workbook.Worksheets.Add(worksheetName);
            for (var i = 0; i < Enum.GetNames<HeaderColumn>().Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = Enum.GetName(typeof(HeaderColumn), i);
            }

            // Apply styling to the headers
            var headerRange = worksheet.Range(1, 1, 1, Enum.GetNames<HeaderColumn>().Length);
            headerRange.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FF228B22"));
            headerRange.Style.Font.Bold = true;

            var dataToSet = Page(result, maxRowPerSheet, pager);

            //2. tól kezdjük a fejléc miatt
            worksheet.Cell(2, 1).InsertData(dataToSet);
            _ = worksheet.Range("A:F").SetAutoFilter();
            _ = worksheet.Columns("A:F").AdjustToContents();
            _ = worksheet.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            _ = worksheet.Style.Border.InsideBorderColor = XLColor.FromHtml("#FF3B444B");
            _ = worksheet.Style.Font.SetFontColor(XLColor.Black);

            worksheet.FirstRow().Height = 30;
            worksheet.FirstRow().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.FirstRow().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.FirstRow().Style.Font.SetFontColor(XLColor.White);
        }

        if (!File.Exists(filePath)) workbook.SaveAs(filePath);
    }
}