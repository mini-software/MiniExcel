namespace MiniExcelLib.Core.Helpers;

internal static class ThrowHelper
{
    internal static void ThrowIfInvalidOpenXml(Stream stream)
    {
        var probe = new byte[8];
        stream.Seek(0, SeekOrigin.Begin);
        var read = stream.Read(probe, 0, probe.Length);
        if (read != probe.Length)
            throw new InvalidDataException("The file/stream does not contain enough data to be processed.");
            
        stream.Seek(0, SeekOrigin.Begin);

        // OpenXml format can be any ZIP archive
        if (probe is not [0x50, 0x4B, ..])
            throw new InvalidDataException("The file is not a valid OpenXml document.");
    }
}
