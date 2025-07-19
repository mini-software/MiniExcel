using System.Collections.Concurrent;

namespace MiniExcelLib.Helpers;

// For Row/Column Index
public static class ColumnHelper
{
    private const int GeneralColumnIndex = 255;
    private const int MaxColumnIndex = 16383;
    
    private static readonly ConcurrentDictionary<int, string> IntMappingAlphabet = new();
    private static readonly ConcurrentDictionary<string, int> AlphabetMappingInt = new();

    private static int _intMappingAlphabetCount;

    static ColumnHelper()
    {
        _intMappingAlphabetCount = IntMappingAlphabet.Count;
        CheckAndSetMaxColumnIndex(GeneralColumnIndex);
    }

    public static string GetAlphabetColumnName(int columnIndex)
    {
        CheckAndSetMaxColumnIndex(columnIndex);
        return IntMappingAlphabet.TryGetValue(columnIndex, out var value) ? value : throw new KeyNotFoundException();
    }

    public static int GetColumnIndex(string? columnName)
    {
        columnName ??= "";
        if (AlphabetMappingInt.TryGetValue(columnName, out var columnIndex))
            CheckAndSetMaxColumnIndex(columnIndex);
        
        return columnIndex;
    }

    private static void CheckAndSetMaxColumnIndex(int columnIndex)
    {
        if (columnIndex < _intMappingAlphabetCount)
            return;
        
        if (columnIndex > MaxColumnIndex)
            throw new InvalidDataException($"ColumnIndex {columnIndex} over excel vaild max index.");
            
        for (int i = IntMappingAlphabet.Count; i <= columnIndex; i++)
        {
            IntMappingAlphabet.AddOrUpdate(i, IntToLetters(i), (_, _) => IntToLetters(i));
            AlphabetMappingInt.AddOrUpdate(IntToLetters(i), i, (_, _) => i);
        }
        _intMappingAlphabetCount = IntMappingAlphabet.Count;
    }

    private static string IntToLetters(int value)
    {
        value++;
        var result = string.Empty;
        while (--value >= 0)
        {
            result = (char)('A' + value % 26) + result;
            value /= 26;
        }
        return result;
    }
}