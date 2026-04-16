using System.IO.Compression;

namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Supplements base classes with synchronous method counterparts, ensuring compatibility with the SyncMethodGenerator
/// by providing missing entry points without requiring manual preprocessor directives (#if SYNC_ONLY)
/// </summary>
public static class SynchronousHelper
{
    extension(ZipArchive)
    {
        public static ZipArchive Create(Stream stream, ZipArchiveMode mode, bool leaveOpen, Encoding? encoding = null) 
            => new(stream, mode, leaveOpen, encoding);
    }
}
