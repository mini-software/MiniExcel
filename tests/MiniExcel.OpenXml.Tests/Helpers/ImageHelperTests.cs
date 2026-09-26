using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.Helpers;

public class ImageHelperTests
{
    private static byte[] TestPng() => File.ReadAllBytes(PathHelper.GetFile("xlsx/Issue327/TestIssue327.png"));

    [Fact]
    public void GetImageSize_ReadsPngHeader()
    {
        var size = ImageHelper.GetImageSize(TestPng());

        Assert.NotNull(size);
        Assert.Equal(1920, size!.Value.Width);
        Assert.Equal(1032, size.Value.Height);
    }

    [Fact]
    public void GetImageSize_ReadsGifHeader()
    {
        byte[] gif = [(byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 100, 0, 50, 0];

        var size = ImageHelper.GetImageSize(gif);

        Assert.NotNull(size);
        Assert.Equal(100, size!.Value.Width);
        Assert.Equal(50, size.Value.Height);
    }

    [Fact]
    public void GetImageSize_ReadsBmpHeader()
    {
        var bmp = new byte[26];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        bmp[14] = 40;   // BITMAPINFOHEADER
        bmp[18] = 100;  // width
        bmp[22] = 50;   // height

        var size = ImageHelper.GetImageSize(bmp);

        Assert.NotNull(size);
        Assert.Equal(100, size!.Value.Width);
        Assert.Equal(50, size.Value.Height);
    }

    [Fact]
    public void GetImageSize_ReadsJpegHeader()
    {
        var jpeg = new byte[20];
        jpeg[0] = 0xFF;
        jpeg[1] = 0xD8;
        jpeg[2] = 0xFF;
        jpeg[3] = 0xE0; // APP0: matches the signature used by GetImageFormat
        jpeg[4] = 0x00;
        jpeg[5] = 0x04; // APP0 segment length
        jpeg[8] = 0xFF;
        jpeg[9] = 0xC0; // SOF0
        jpeg[10] = 0x00;
        jpeg[11] = 0x11; // segment length
        jpeg[12] = 0x08; // sample precision
        jpeg[13] = 0x00;
        jpeg[14] = 50;   // height
        jpeg[15] = 0x00;
        jpeg[16] = 100;  // width

        var size = ImageHelper.GetImageSize(jpeg);

        Assert.NotNull(size);
        Assert.Equal(100, size!.Value.Width);
        Assert.Equal(50, size.Value.Height);
    }

    [Fact]
    public void GetImageSize_ReadsTiffHeader()
    {
        var tiff = new byte[34];
        tiff[0] = (byte)'I';
        tiff[1] = (byte)'I';
        tiff[2] = 0x2A;
        tiff[3] = 0x00;
        tiff[4] = 0x08; // offset to the first IFD
        tiff[8] = 0x02; // two entries
        tiff[10] = 0x00;
        tiff[11] = 0x01; // tag 256: image width
        tiff[12] = 0x03; // SHORT
        tiff[14] = 0x01; // count
        tiff[18] = 0x64; // 100
        tiff[22] = 0x01;
        tiff[23] = 0x01; // tag 257: image height
        tiff[24] = 0x03; // SHORT
        tiff[26] = 0x01; // count
        tiff[30] = 0x32; // 50

        var size = ImageHelper.GetImageSize(tiff);

        Assert.NotNull(size);
        Assert.Equal(100, size!.Value.Width);
        Assert.Equal(50, size.Value.Height);
    }

    [Fact]
    public void GetImageSize_ReturnsNullForUnknownData()
        => Assert.Null(ImageHelper.GetImageSize([1, 2, 3, 4, 5, 6, 7, 8]));

    [Fact]
    public void GetImageSize_ReturnsNullForNullBytes()
        => Assert.Null(ImageHelper.GetImageSize(null));
}
