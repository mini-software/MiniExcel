using MiniExcelLibs.Utils;
using System;
using System.Collections;
using System.Data;

namespace MiniExcelLibs.WriteAdapter
{
    internal static class MiniExcelWriteAdapterFactory
    {
        public static bool TryGetAsyncWriteAdapter(object values, Configuration configuration, out IAsyncMiniExcelWriteAdapter writeAdapter)
        {
            writeAdapter = null;
            if (values.GetType().IsAsyncEnumerable(out var genericArgument))
            {
                var writeAdapterType = typeof(AsyncEnumerableWriteAdapter<>).MakeGenericType(genericArgument);
                writeAdapter = (IAsyncMiniExcelWriteAdapter)Activator.CreateInstance(writeAdapterType, values, configuration);
                return true;
            }
            if (values is IMiniExcelDataReader miniExcelDataReader)
            {
                writeAdapter = new MiniExcelDataReaderWriteAdapter(miniExcelDataReader, configuration);
                return true;
            }

            return false;
        }

        public static IMiniExcelWriteAdapter GetWriteAdapter(object values, Configuration configuration)
        {
            switch (values)
            {
                case IDataReader dataReader:
                    return new DataReaderWriteAdapter(dataReader, configuration);
                case IEnumerable enumerable:
                    return new EnumerableWriteAdapter(enumerable, configuration);
                case DataTable dataTable:
                    return new DataTableWriteAdapter(dataTable, configuration);
                default:
                    throw new NotImplementedException();
            }
        }
    }
}
