using System.Collections.ObjectModel;

namespace MiniExcelLib.OpenXml.Zip;

/// Copy & modified by ExcelDataReader ZipWorker @MIT License
internal sealed partial class OpenXmlZip : IDisposable, IAsyncDisposable
{
    private static readonly XmlReaderSettings XmlSettings = new()
    {
        IgnoreComments = true,
        IgnoreWhitespace = true,
        XmlResolver = null,
    };

    private bool _disposed;

    internal ZipArchive ZipFile { get; }
    internal Dictionary<string, ZipArchiveEntry> Entries { get; }
    internal ReadOnlyCollection<ZipArchiveEntry> EntryCollection { get; set; } = new([]);
    

    private OpenXmlZip(ZipArchive zipArchive,  Dictionary<string, ZipArchiveEntry> entries)
    {
        ZipFile = zipArchive;
        Entries = entries;
    }

    // todo: convert to ValueTask and create auxiliary methods to avoid generation of async state machine for framework versions lower than .NET 10
    [CreateSyncVersion]
    internal static async Task<OpenXmlZip> CreateAsync(Stream fileStream, ZipArchiveMode mode = ZipArchiveMode.Read, bool leaveOpen = false, Encoding? entryNameEncoding = null, bool isUpdateMode = true, CancellationToken cancellationToken = default)
    {
        entryNameEncoding ??= Encoding.UTF8;
#if NET10_0_OR_GREATER
        var zipFile = await ZipArchive.CreateAsync(fileStream, mode, leaveOpen, entryNameEncoding, cancellationToken).ConfigureAwait(false);
#else
        var zipFile = new ZipArchive(fileStream, mode, leaveOpen, entryNameEncoding);
#endif

        if (!isUpdateMode)
            return new OpenXmlZip(zipFile, []);

        try
        {
            var entries = zipFile.Entries.ToDictionary(entry => entry.FullName.Replace('\\', '/'), entry => entry, StringComparer.OrdinalIgnoreCase);
            return new OpenXmlZip(zipFile, entries)
            {
                EntryCollection = zipFile.Entries
            };
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidDataException("The file doesn't contain valid OpenXml data.", ex);
        }
    }

    public ZipArchiveEntry? GetEntry(string path) => Entries.GetValueOrDefault(path);

    public XmlReader? GetXmlReader(string path) => GetEntry(path) is { } entry 
        ? XmlReader.Create(entry.Open(), XmlSettings)
        : null;


    public void Dispose()
    {
        if (!_disposed)
        {
            ZipFile.Dispose();
            _disposed = true;
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (!_disposed)
        {
#if NET10_0_OR_GREATER
            await ZipFile.DisposeAsync().ConfigureAwait(false);
#else
            ZipFile.Dispose();
#endif
            _disposed = true;
        }
    }
}
