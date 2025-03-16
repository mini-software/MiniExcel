namespace MiniExcelLibs.OpenXml
{
    public class SheetInfo
    {
        public SheetInfo(uint id, uint index, string name, SheetState sheetState, bool active)
        {
            Id = id;
            Index = index;
            Name = name;
            State = sheetState;
            Active = active;
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
        /// <summary>
        /// Indicates whether the sheet is active
        /// </summary>
        public bool Active { get; }
    }

    public enum SheetState { Visible, Hidden, VeryHidden }
}
