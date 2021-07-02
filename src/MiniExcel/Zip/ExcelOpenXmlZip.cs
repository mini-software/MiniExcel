using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace MiniExcelLibs.Zip
{
    /// Copy & modified by ExcelDataReader ZipWorker @MIT License
    internal class ExcelOpenXmlZip : IDisposable
    {
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private bool _disposed;
        private Stream _zipStream;
        internal MiniExcelZipArchive zipFile;
        public ReadOnlyCollection<ZipArchiveEntry> entries;

       private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings
       {
           IgnoreComments = true,
           IgnoreWhitespace = true,
           XmlResolver = null,
       };
        public ExcelOpenXmlZip(Stream fileStream, ZipArchiveMode mode = ZipArchiveMode.Read, bool leaveOpen = false, Encoding entryNameEncoding = null)
        {
            _zipStream = fileStream ?? throw new ArgumentNullException(nameof(fileStream));
            zipFile = new MiniExcelZipArchive(fileStream, mode, leaveOpen, entryNameEncoding);
            _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            try
            {
                entries = zipFile.Entries; //TODO:need to remove
            }
            catch (InvalidDataException e)
            {
                throw new InvalidDataException($"It's not legal excel zip, please check or issue for me. {e.Message}");
            }

            foreach (var entry in zipFile.Entries)
            {
                _entries.Add(entry.FullName.Replace('\\', '/'), entry);
            }
        }

        public ZipArchiveEntry GetEntry(string path)
        {
            if (_entries.TryGetValue(path, out var entry))
                return entry;
            return null;
        }

        public XmlReader GetXmlReader(string path)
        {
            var entry = GetEntry(path);
            if (entry != null)
                return XmlReader.Create(entry.Open(), XmlSettings);
            return null;
        }

        ~ExcelOpenXmlZip()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!_disposed)
            {
                if (disposing)
                {
                    if (zipFile != null)
                    {
                        zipFile.Dispose();
                        zipFile = null;
                    }

                    if (_zipStream != null)
                    {
                        _zipStream.Dispose();
                        _zipStream = null;
                    }
                }

                _disposed = true;
            }
        }
    }
}
