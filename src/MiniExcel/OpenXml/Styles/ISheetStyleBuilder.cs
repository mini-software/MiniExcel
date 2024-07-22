using MiniExcelLibs.Attributes;
using System.Collections.Generic;

namespace MiniExcelLibs.OpenXml.Styles {
    public interface ISheetStyleBuilder
    {
        string Build( ICollection<ExcelColumnAttribute> columns );
    }

}
