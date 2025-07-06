using System.Collections.ObjectModel;

namespace MiniExcelLib.Core.OpenXml.Zip;

/// Copy & modified by ExcelDataReader ZipWorker @MIT License
internal class OpenXmlZip : IDisposable
{
    private bool _disposed;

    public ReadOnlyCollection<ZipArchiveEntry> EntryCollection = new([]);
    
    internal readonly Dictionary<string, ZipArchiveEntry> Entries;
    internal MiniExcelZipArchive? ZipFile;

    private static readonly XmlReaderSettings XmlSettings = new()
    {
        IgnoreComments = true,
        IgnoreWhitespace = true,
        XmlResolver = null,
    };
    
    public OpenXmlZip(Stream fileStream, ZipArchiveMode mode = ZipArchiveMode.Read, bool leaveOpen = false, Encoding? entryNameEncoding = null, bool isUpdateMode = true)
    {
        entryNameEncoding ??= Encoding.UTF8;
        ZipFile = new MiniExcelZipArchive(fileStream, mode, leaveOpen, entryNameEncoding);
        Entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        
        if (!isUpdateMode)
            return;
        
        try
        {
            EntryCollection = ZipFile.Entries; //TODO:need to remove
        }
        catch (InvalidDataException e)
        {
            throw new InvalidDataException($"The file doesn't contain valid OpenXml, please check the project issues or open one. {e.Message}");
        }

        foreach (var entry in ZipFile.Entries)
        {
            Entries.Add(entry.FullName.Replace('\\', '/'), entry);
        }
    }

    public ZipArchiveEntry? GetEntry(string path)
    {
        return Entries.TryGetValue(path, out var entry) ? entry : null;
    }

    public XmlReader? GetXmlReader(string path)
    {
        var entry = GetEntry(path);
        return entry is not null ? XmlReader.Create(entry.Open(), XmlSettings) : null;
    }

    ~OpenXmlZip()
    {
        Dispose(false);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        // Check to see if Dispose has already been called.
        if (!_disposed)
        {
            if (disposing)
            {
                if (ZipFile is not null)
                {
                    ZipFile.Dispose();
                    ZipFile = null;
                }
            }

            _disposed = true;
        }
    }
}