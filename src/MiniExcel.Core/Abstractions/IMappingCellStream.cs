namespace MiniExcelLib.Core.Abstractions;

public interface IMappingCellStream
{
    IMiniExcelWriteAdapter CreateAdapter();
}
