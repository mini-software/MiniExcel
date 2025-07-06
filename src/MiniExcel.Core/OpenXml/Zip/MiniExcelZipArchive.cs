namespace MiniExcelLib.Core.OpenXml.Zip;

public class MiniExcelZipArchive(Stream stream, ZipArchiveMode mode, bool leaveOpen, Encoding entryNameEncoding)
    : ZipArchive(stream, mode, leaveOpen, entryNameEncoding)
{
    public new void Dispose()
    {
        base.Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}