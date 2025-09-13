namespace MiniExcelLib.Core.Helpers;

internal static class ThrowHelper
{
    private static readonly byte[] ZipArchiveHeader = [0x50, 0x4B];
    private const int ExcelMaxSheetNameLength = 31;
    
    internal static void ThrowIfInvalidOpenXml(Stream stream)
    {
        var probe = new byte[8];
        stream.Seek(0, SeekOrigin.Begin);
        var read = stream.Read(probe, 0, probe.Length);
        if (read != probe.Length)
            throw new InvalidDataException("The file/stream does not contain enough data to be processed.");
            
        stream.Seek(0, SeekOrigin.Begin);

        // OpenXml format can be any ZIP archive
        if (!probe.StartsWith(ZipArchiveHeader))
            throw new InvalidDataException("The file is not a valid OpenXml document.");
    }

    internal static void ThrowIfInvalidSheetName(string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("Sheet names cannot be empty or null");

        if (sheetName.Length > ExcelMaxSheetNameLength)
            throw new ArgumentException("Sheet names must be less than 31 characters");
    }
}
