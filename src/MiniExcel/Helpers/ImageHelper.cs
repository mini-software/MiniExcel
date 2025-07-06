namespace MiniExcelLib.Helpers;

internal static class ImageHelper
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
}