using System;

namespace MiniExcelLibs.OpenXml
{
    internal sealed class SheetRecord
    {
        public SheetRecord(string name, string state, uint id, string rid, bool active)
        {
            Name = name;
            State = state;
            Id = id;
            Rid = rid;
            Active = active;
        }

        public string Name { get; }

        public string State { get; set; }

        public uint Id { get; }

        public string Rid { get; set; }

        public string Path { get; set; }

        public bool Active { get; }

        public SheetInfo ToSheetInfo(uint index)
        {
            if (string.IsNullOrEmpty(State))
            {
                return new SheetInfo(Id, index, Name, SheetState.Visible, Active);
            }
            if (Enum.TryParse(State, true, out SheetState stateEnum))
            {
                return new SheetInfo(Id, index, Name, stateEnum, Active);
            }
            throw new ArgumentException($"Unable to parse sheet state. Sheet name: {Name}");
        }
    }
}
