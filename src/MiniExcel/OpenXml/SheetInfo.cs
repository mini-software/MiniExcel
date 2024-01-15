namespace MiniExcelLibs.OpenXml
{
    public class SheetInfo
    {
        public SheetInfo(uint id, uint index, string name, SheetState sheetState)
        {
            Id = id;
            Index = index;
            Name = name;
            State = sheetState;
        }

        /// <summary>
        /// Internal sheet id - depends on the order in which the sheet is added
        /// </summary>
        public uint Id { get; }
        /// <summary>
        /// Next sheet index - numbered from 0
        /// </summary>
        public uint Index { get; }
        /// <summary>
        /// Sheet name
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// Sheet visibility state
        /// </summary>
        public SheetState State { get; }
    }

    public enum SheetState { Visible, Hidden, VeryHidden }
}
