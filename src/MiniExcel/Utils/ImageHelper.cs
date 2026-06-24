namespace MiniExcelLibs.Utils;

internal static class ImageHelper
{
#if NET
    private static ReadOnlySpan<byte> Bmp => "BM"u8;                  // BMP
    private static ReadOnlySpan<byte> Gif => "GIF"u8;                 // GIF
    private static ReadOnlySpan<byte> Png => [137, 80, 78, 71];       // PNG
    private static ReadOnlySpan<byte> Tiff => "II*"u8;                // TIFF
    private static ReadOnlySpan<byte> Tiff2 => "MM*"u8;               // TIFF
    private static ReadOnlySpan<byte> Jpeg => [255, 216, 255, 224];   // JPEG
    private static ReadOnlySpan<byte> Jpeg2 => [255, 216, 255, 225];  // JPEG canon
#else
    private static readonly byte[] Bmp = [(byte)'B', (byte)'M'];            // BMP
    private static readonly byte[] Gif = [(byte)'G', (byte)'I', (byte)'F']; // GIF
    private static readonly byte[] Png = [137, 80, 78, 71];                 // PNG
    private static readonly byte[] Tiff = [(byte)'I', (byte)'I'];           // TIFF
    private static readonly byte[] Tiff2 = [(byte)'M', (byte)'M'];          // TIFF
    private static readonly byte[] Jpeg = [255, 216, 255, 224];             // JPEG
    private static readonly byte[] Jpeg2 = [255, 216, 255, 225];            // JPEG canon
#endif

    public static ImageFormat GetImageFormat(
#if NET
        ReadOnlySpan<byte> bytes
#else
        byte[] bytes
#endif
    )
    {
        if (bytes.StartsWith(Bmp))
            return ImageFormat.bmp;

        if (bytes.StartsWith(Gif))
            return ImageFormat.gif;

        if (bytes.StartsWith(Png))
            return ImageFormat.png;

        if (bytes.StartsWith(Tiff) || bytes.StartsWith(Tiff2))
            return ImageFormat.tiff;

        if (bytes.StartsWith(Jpeg) || bytes.StartsWith(Jpeg2))
            return ImageFormat.jpg;

        return ImageFormat.unknown;
    }
    
    public enum ImageFormat
    {
        bmp,
        jpg,
        gif,
        tiff,
        png,
        unknown
    }
}
