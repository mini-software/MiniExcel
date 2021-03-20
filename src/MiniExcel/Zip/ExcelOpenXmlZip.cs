using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace MiniExcelLibs.Zip
{
    /// Copy & modified by ExcelDataReader ZipWorker
    internal class ExcelOpenXmlZip : IDisposable
    {
	   private readonly Dictionary<string, ZipArchiveEntry> _entries;
	   private bool _disposed;
	   private Stream _zipStream;
	   private ZipArchive _zipFile;
	   public ReadOnlyCollection<ZipArchiveEntry> Entries;
	   private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings
	   {
		  IgnoreComments = true,
		  IgnoreWhitespace = true,
		  XmlResolver = null,
	   };
	   public ExcelOpenXmlZip(Stream fileStream)
	   {
		  _zipStream = fileStream ?? throw new ArgumentNullException(nameof(fileStream));
		  _zipFile = new ZipArchive(fileStream);
		  _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
		  Entries = _zipFile.Entries; //TODO:need to remove
		  foreach (var entry in _zipFile.Entries)
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
				if (_zipFile != null)
				{
				    _zipFile.Dispose();
				    _zipFile = null;
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
