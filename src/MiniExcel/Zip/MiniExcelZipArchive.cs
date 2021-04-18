using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace MiniExcelLibs.Zip
{
    public class MiniExcelZipArchive : ZipArchive
    {
	   public MiniExcelZipArchive(Stream stream, ZipArchiveMode mode, bool leaveOpen, Encoding entryNameEncoding)
		  : base(stream, mode, leaveOpen, entryNameEncoding)
	   {
	   }

	   public new void Dispose()
	   {
		  Dispose(disposing: true);
		  GC.SuppressFinalize(this);
	   }
    }
}
