namespace MiniExcel
{
    internal class ZipPackageInfo
    {
        public string Xml { get; set; }
        public string ContentType { get; set; }
        //public CompressionOption CompressionOption { get; set; } = CompressionOption.Normal;
        public ZipPackageInfo(string xml, string contentType)
        {
            Xml = xml;
            ContentType = contentType;
        }
    }
}
