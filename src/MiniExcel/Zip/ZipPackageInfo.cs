namespace MiniExcelLibs.Zip
{
    internal class ZipPackageInfo
    {
        public string Xml { get; set; }
        public string ContentType { get; set; }
        public ZipPackageInfo(string xml, string contentType)
        {
            Xml = xml;
            ContentType = contentType;
        }
    }
}
