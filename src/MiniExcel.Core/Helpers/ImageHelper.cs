namespace MiniExcelLib.Core.Helpers;

public static class ImageHelper
{
    private static readonly byte[] Bmp = [(byte)'B', (byte)'M'];
    private static readonly byte[] Gif = [(byte)'G', (byte)'I', (byte)'F'];
    private static readonly byte[] Png = [137, 80, 78, 71];
    private static readonly byte[] Tiff = [73, 73, 42];
    private static readonly byte[] Tiff2 = [77, 77, 42];
    private static readonly byte[] Jpeg = [255, 216, 255, 224];
    private static readonly byte[] Jpeg2 = [255, 216, 255, 225];
    
    public enum ImageFormat
    {
        Bmp,
        Jpg,
        Gif,
        Tiff,
        Png,
        Unknown
    }

    public static ImageFormat GetImageFormat(byte[] bytes)
    {
        if (bytes.StartsWith(Bmp))
            return ImageFormat.Bmp;
        
        if (bytes.StartsWith(Gif))
            return ImageFormat.Gif;

        if (bytes.StartsWith(Png))
            return ImageFormat.Png;

        if (bytes.StartsWith(Tiff) || bytes.StartsWith(Tiff2))
            return ImageFormat.Tiff;

        if (bytes.StartsWith(Jpeg) || bytes.StartsWith(Jpeg2))
            return ImageFormat.Jpg;

        return ImageFormat.Unknown;
    }

    /// <summary>
    /// Reads the pixel dimensions of an image from its header. Returns <c>null</c> when the format is
    /// not recognised or the header is truncated.
    /// </summary>
    public static (int Width, int Height)? GetImageSize(byte[]? bytes)
    {
        if (bytes is null || bytes.Length < 8)
            return null;

        if (bytes.StartsWith(Png))
            return GetPngSize(bytes);

        if (bytes.StartsWith(Gif))
            return GetGifSize(bytes);

        if (bytes.StartsWith(Bmp))
            return GetBmpSize(bytes);

        if (bytes.StartsWith(Jpeg) || bytes.StartsWith(Jpeg2))
            return GetJpegSize(bytes);

        if (bytes.StartsWith(Tiff) || bytes.StartsWith(Tiff2))
            return GetTiffSize(bytes);

        return null;
    }

    private static (int, int)? GetPngSize(byte[] bytes)
    {
        // 8-byte signature, 4-byte chunk length, then the "IHDR" chunk carrying width and height as
        // big-endian 32-bit integers.
        if (bytes.Length < 24 || bytes[12] != 'I' || bytes[13] != 'H' || bytes[14] != 'D' || bytes[15] != 'R')
            return null;

        var width = ReadInt32BigEndian(bytes, 16);
        var height = ReadInt32BigEndian(bytes, 20);
        return width > 0 && height > 0 ? (width, height) : null;
    }

    private static (int, int)? GetGifSize(byte[] bytes)
    {
        // Logical screen descriptor: width and height as little-endian 16-bit integers.
        if (bytes.Length < 10)
            return null;

        var width = bytes[6] | (bytes[7] << 8);
        var height = bytes[8] | (bytes[9] << 8);
        return width > 0 && height > 0 ? (width, height) : null;
    }

    private static (int, int)? GetBmpSize(byte[] bytes)
    {
        if (bytes.Length < 26)
            return null;

        // A BITMAPCOREHEADER stores 16-bit dimensions; the more common BITMAPINFOHEADER family uses
        // 32-bit ones, with a negative height meaning a top-down bitmap.
        if (ReadInt32LittleEndian(bytes, 14) == 12)
        {
            var coreWidth = bytes[18] | (bytes[19] << 8);
            var coreHeight = bytes[20] | (bytes[21] << 8);
            return coreWidth > 0 && coreHeight > 0 ? (coreWidth, coreHeight) : null;
        }

        var width = ReadInt32LittleEndian(bytes, 18);
        var height = Math.Abs((long)ReadInt32LittleEndian(bytes, 22));
        return width > 0 && height is > 0 and <= int.MaxValue ? (width, (int)height) : null;
    }

    private static (int, int)? GetJpegSize(byte[] bytes)
    {
        var index = 2;
        while (index + 8 < bytes.Length)
        {
            if (bytes[index] != 0xFF)
            {
                index++;
                continue;
            }

            var marker = bytes[index + 1];
            if (marker == 0xFF)
            {
                index++;
                continue;
            }

            // Standalone markers (RSTn, SOI, EOI, TEM) have no payload.
            if (marker == 0x01 || marker is >= 0xD0 and <= 0xD9)
            {
                index += 2;
                continue;
            }

            // Start of scan: any frame header would have been found before this point.
            if (marker == 0xDA)
                break;

            var segmentLength = (bytes[index + 2] << 8) | bytes[index + 3];
            if (segmentLength < 2)
                break;

            // SOF0..SOF15, excluding DHT (C4), JPG (C8) and DAC (CC).
            var isFrameHeader = marker is >= 0xC0 and <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
            if (isFrameHeader)
            {
                var height = (bytes[index + 5] << 8) | bytes[index + 6];
                var width = (bytes[index + 7] << 8) | bytes[index + 8];
                return width > 0 && height > 0 ? (width, height) : null;
            }

            index += 2 + segmentLength;
        }

        return null;
    }

    private static (int, int)? GetTiffSize(byte[] bytes)
    {
        if (bytes.Length < 8)
            return null;

        var littleEndian = bytes[0] == 'I';
        var ifdOffset = ReadInt32(bytes, 4, littleEndian);
        if (ifdOffset < 8 || ifdOffset + 2 > bytes.Length)
            return null;

        var entryCount = ReadUInt16(bytes, ifdOffset, littleEndian);
        int? width = null;
        int? height = null;

        for (var i = 0; i < entryCount; i++)
        {
            var entryOffset = ifdOffset + 2 + (i * 12);
            if (entryOffset + 12 > bytes.Length)
                break;

            var tag = ReadUInt16(bytes, entryOffset, littleEndian);
            if (tag != 256 && tag != 257)
                continue;

            var fieldType = ReadUInt16(bytes, entryOffset + 2, littleEndian);
            int value;
            if (fieldType == 3) // SHORT
                value = ReadUInt16(bytes, entryOffset + 8, littleEndian);
            else if (fieldType == 4) // LONG
                value = ReadInt32(bytes, entryOffset + 8, littleEndian);
            else
                continue;

            if (tag == 256)
                width = value;
            else
                height = value;
        }

        return width is > 0 && height is > 0 ? (width.Value, height.Value) : null;
    }

    private static int ReadInt32BigEndian(byte[] bytes, int offset)
        => (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static int ReadInt32LittleEndian(byte[] bytes, int offset)
        => bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    private static int ReadInt32(byte[] bytes, int offset, bool littleEndian)
        => littleEndian ? ReadInt32LittleEndian(bytes, offset) : ReadInt32BigEndian(bytes, offset);

    private static int ReadUInt16(byte[] bytes, int offset, bool littleEndian)
        => littleEndian
            ? bytes[offset] | (bytes[offset + 1] << 8)
            : (bytes[offset] << 8) | bytes[offset + 1];
}
