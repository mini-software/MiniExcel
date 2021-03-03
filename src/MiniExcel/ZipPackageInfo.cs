namespace MiniExcel
{
    using System.IO.Packaging;

    internal class ZipPackageInfo
    {
        public string Xml { get; set; }
        public string ContentType { get; set; }
        public CompressionOption CompressionOption { get; set; } = CompressionOption.Normal;
        public ZipPackageInfo(string xml, string contentType, CompressionOption CompressionOption = CompressionOption.Normal)
        {
            Xml = xml;
            ContentType = contentType;
        }
    }
}
