using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace MiniExcelLibs.Zip;

public class MiniExcelZipArchive(Stream stream, ZipArchiveMode mode, bool leaveOpen, Encoding entryNameEncoding)
    : ZipArchive(stream, mode, leaveOpen, entryNameEncoding)
{
    public new void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}