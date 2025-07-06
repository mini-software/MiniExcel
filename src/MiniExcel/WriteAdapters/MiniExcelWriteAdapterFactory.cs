using MiniExcelLib.Abstractions;
using MiniExcelLib.Helpers;

namespace MiniExcelLib.WriteAdapters;

public static class MiniExcelWriteAdapterFactory
{
    public static bool TryGetAsyncWriteAdapter(object values, MiniExcelBaseConfiguration configuration, out IMiniExcelWriteAdapterAsync? writeAdapter)
    {
        writeAdapter = null;
        if (values.GetType().IsAsyncEnumerable(out var genericArgument))
        {
            var writeAdapterType = typeof(AsyncEnumerableWriteAdapter<>).MakeGenericType(genericArgument);
            writeAdapter = (IMiniExcelWriteAdapterAsync)Activator.CreateInstance(writeAdapterType, values, configuration);
            return true;
        }
        
        if (values is IMiniExcelDataReader miniExcelDataReader)
        {
            writeAdapter = new MiniExcelDataReaderWriteAdapter(miniExcelDataReader, configuration);
            return true;
        }

        return false;
    }

    public static IMiniExcelWriteAdapter GetWriteAdapter(object values, MiniExcelBaseConfiguration configuration)
    {
        return values switch
        {
            IDataReader dataReader => new DataReaderWriteAdapter(dataReader, configuration),
            IEnumerable enumerable => new EnumerableWriteAdapter(enumerable, configuration),
            DataTable dataTable => new DataTableWriteAdapter(dataTable, configuration),
            _ => throw new NotImplementedException()
        };
    }
}