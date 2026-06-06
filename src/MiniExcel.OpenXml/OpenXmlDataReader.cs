using MiniExcelLib.OpenXml.Reader;

namespace MiniExcelLib.OpenXml;

public sealed class OpenXmlDataReader : MiniExcelDataReaderBase
{
    private readonly bool _expectsSingleResult;

    private readonly string[] _sheetNames;
    private int _currentSheetIndex;
    private string _currentSheetName;
    
    private string _startCell;


    private OpenXmlDataReader(OpenXmlReader reader, string sheetName, bool hasHeaderRow, string startCell, bool isAsyncSource, bool expectsSingleResult, string[] sheetNames, OpenXmlConfiguration configuration)
        : base(reader, hasHeaderRow, isAsyncSource, configuration)
    {
        _startCell = startCell;
        _expectsSingleResult = expectsSingleResult;
        _sheetNames =  sheetNames;
        _currentSheetName = sheetName;
    }

    internal static OpenXmlDataReader Create(Stream stream, bool hasHeaderRow = false, string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, bool leaveOpen = false)
    {
        OpenXmlReader? reader = null;
        OpenXmlDataReader? dataReader = null;
        configuration ??= OpenXmlConfiguration.Default;

        try
        {
            reader = OpenXmlReader.Create(stream, configuration, leaveOpen);

            bool isSingleResult = false;
            string[] sheetNames = [];
            if (string.IsNullOrEmpty(sheetName))
            {
                var sheets = OpenXmlReader.GetWorkbookRels(reader.Archive.EntryCollection);
                sheetNames = sheets?.Select(s => s.Name).ToArray() ?? [];
            }
            else
            {
                isSingleResult = true;
            }

            dataReader = new OpenXmlDataReader(reader, sheetName ?? sheetNames[0], hasHeaderRow, startCell, isAsyncSource: false, isSingleResult, sheetNames, configuration)
            {
                Source = reader.Query(hasHeaderRow, sheetName, startCell).GetEnumerator(),
            };

            if (dataReader.Source!.MoveNext())
            {
                dataReader.Columns = dataReader.Source.Current?.Keys.ToList() ?? [];
                dataReader.FieldCount = dataReader.Columns.Count;
            }
            else
            {
                dataReader.IsEmpty = true;
            }

            var result = dataReader;
            dataReader = null;
            reader = null;
            stream = null!;
            
            return result;
        }
        finally
        {
            dataReader?.Dispose();
            reader?.Dispose();
            
            if (!leaveOpen)
                ((Stream?)stream)?.Dispose();
        }
    }

    internal static async Task<OpenXmlDataReader> CreateAsync(Stream stream, bool hasHeaderRow = false, string? sheetName = null, string startCell = "A1", OpenXmlConfiguration? configuration = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        OpenXmlReader? reader = null;
        OpenXmlDataReader? dataReader = null;
        configuration ??= OpenXmlConfiguration.Default;

        try
        {
            reader = await OpenXmlReader.CreateAsync(stream, configuration, leaveOpen, cancellationToken).ConfigureAwait(false);

            bool isSingleResult = false;
            string[] sheetNames = [];
            if (string.IsNullOrEmpty(sheetName))
            {
                var sheets = await OpenXmlReader.GetWorkbookRelsAsync(reader.Archive.EntryCollection, cancellationToken).ConfigureAwait(false);
                sheetNames = sheets?.Select(s => s.Name).ToArray() ?? [];
            }
            else
            {
                isSingleResult = true;
            }

            dataReader = new OpenXmlDataReader(reader, sheetName ?? sheetNames[0], hasHeaderRow, startCell, isAsyncSource: true, isSingleResult, sheetNames, configuration)
            {
                AsyncSource = reader.QueryAsync(hasHeaderRow, sheetName, startCell, cancellationToken).GetAsyncEnumerator(cancellationToken),
            };

            if (await dataReader.AsyncSource.MoveNextAsync().ConfigureAwait(false))
            {
                dataReader.Columns = dataReader.AsyncSource.Current?.Keys.ToList() ?? [];
                dataReader.FieldCount = dataReader.Columns.Count;
            }
            else
            {
                dataReader.IsEmpty = true;
            }

            var result = dataReader;
            dataReader = null;
            reader = null;
            stream = null!;
            
            return result;
        }
        finally
        {
            if (dataReader is not null)
                await dataReader.DisposeAsync().ConfigureAwait(false);
            
            if (reader?.DisposeAsync() is { } disposeTask)
                await disposeTask.ConfigureAwait(false);
            
            if (!leaveOpen && (Stream?)stream is not null)
                await stream.DisposeAsync().ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Returns the name of the worksheet currently being read. 
    /// </summary>
    public string GetWorksheetName() => _currentSheetName;
    
    private void NextResultCore()
    {
        Source!.Dispose();
        Source = MiniExcelReader.Query(HasHeaderRow, _sheetNames[_currentSheetIndex], _startCell).GetEnumerator();
 
        if (Source!.MoveNext())
        {
            Columns = Source.Current?.Keys.ToList() ?? [];
            FieldCount = Columns.Count;
        }
        else
        {
            IsEmpty = true;
        }
    }

    public override bool NextResult()
    {
        if (IsClosed)
            throw new InvalidOperationException("The data reader has been closed");

        if (IsAsyncSource)
            throw new InvalidOperationException("The data reader was configured to execute asynchronously");

        if (_expectsSingleResult || _currentSheetIndex + 1 >= _sheetNames.Length)
            return false;

        Schema = null;
        Ordinals.Clear();

        _currentSheetIndex++;
        _currentSheetName = _sheetNames[_currentSheetIndex];
        
        NextResultCore();
        return true;
    }

    public override async Task<bool> NextResultAsync(CancellationToken cancellationToken = default)
    {
        if (IsClosed)
            throw new InvalidOperationException("The data reader has been closed");

        if (_expectsSingleResult || _currentSheetIndex + 1 >= _sheetNames.Length)
            return false;

        Schema = null;
        Ordinals.Clear();

        _currentSheetIndex++;
        _currentSheetName = _sheetNames[_currentSheetIndex];

        if (!IsAsyncSource)
        {
            await Task.Run(NextResultCore, cancellationToken).ConfigureAwait(false);
            return true;
        }

        await AsyncSource!.DisposeAsync().ConfigureAwait(false);
        AsyncSource = MiniExcelReader.QueryAsync(HasHeaderRow, _sheetNames[_currentSheetIndex], _startCell, cancellationToken).GetAsyncEnumerator(cancellationToken);

        if (await AsyncSource!.MoveNextAsync().ConfigureAwait(false))
        {
            Columns = AsyncSource.Current?.Keys.ToList() ?? [];
            FieldCount = Columns.Count;
        }
        else
        {
            IsEmpty = true;
        }

        return true;
    }
}
