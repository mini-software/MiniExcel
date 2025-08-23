namespace MiniExcelLib.Core.Mapping;

internal interface IMappingCellStream
{
    IMiniExcelWriteAdapter CreateAdapter();
}