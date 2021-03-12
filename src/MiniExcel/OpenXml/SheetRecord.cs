namespace MiniExcelLibs.OpenXml
{
    internal sealed class SheetRecord
    {
        public SheetRecord(string name, uint id, string rid)
        {
            Name = name;
            Id = id;
            Rid = rid;
        }

        public string Name { get; }

        public uint Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }
    }
}
