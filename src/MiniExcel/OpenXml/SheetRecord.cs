using System;

namespace MiniExcelLibs.OpenXml
{
    internal sealed class SheetRecord
    {
        public SheetRecord(string name, string state, uint id, string rid)
        {
            Name = name;
            State = state;
            Id = id;
            Rid = rid;
        }

        public string Name { get; }

        public string State { get; set; }

        public uint Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }

        public SheetInfo ToSheetInfo(uint index)
        {
            if (string.IsNullOrEmpty(State))
            {
                return new SheetInfo(Id, index, Name, SheetState.Visible);
            }
            if (Enum.TryParse(State, true, out SheetState stateEnum))
            {
                return new SheetInfo(Id, index, Name, stateEnum);
            }
            throw new ArgumentException($"Unable to parse sheet state. Sheet name: {Name}");
        }
    }
}
