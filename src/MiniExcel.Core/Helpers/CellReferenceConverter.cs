using System.Collections.Concurrent;

namespace MiniExcelLib.Core.Helpers;

public static class CellReferenceConverter
{
    private const int GeneralColumnIndex = 255;
    private const int MaxColumnIndex = 16383;

    private static readonly ConcurrentDictionary<int, string> IntMappingToAlphabet = new();
    private static readonly ConcurrentDictionary<string, int> AlphabetMappingToInt = new();

    static CellReferenceConverter()
    {
        EnsureMappingsUpTo(GeneralColumnIndex);
    }

    public static string GetAlphabeticalIndex(int columnIndex)
    {
        EnsureMappingsUpTo(columnIndex);
        return IntMappingToAlphabet.TryGetValue(columnIndex, out var value) ? value : throw new KeyNotFoundException();
    }

    public static int GetNumericalIndex(string? columnName)
    {
        columnName ??= string.Empty;
        if (AlphabetMappingToInt.TryGetValue(columnName, out var columnIndex))
            EnsureMappingsUpTo(columnIndex);

        return columnIndex;
    }

    private static void EnsureMappingsUpTo(int columnIndex)
    {
        if (columnIndex < IntMappingToAlphabet.Count)
            return;

        if (columnIndex > MaxColumnIndex)
            throw new InvalidDataException($"Column index {columnIndex} exceeds Excel's maximum valid index.");

        for (int i = IntMappingToAlphabet.Count; i <= columnIndex; i++)
        {
            var name = GetColumnFromIndex(i);
            
            IntMappingToAlphabet.TryAdd(i, name);
            AlphabetMappingToInt.TryAdd(name, i);
        }
    }

    private static string GetColumnFromIndex(int value)
    {
        var result = string.Empty;
        
        do
        {
            result = (char)('A' + value % 26) + result;
            value = value / 26 - 1;
        }
        while (value >= 0);
        
        return result;
    }
    
    /// <summary>eg. A2=(1,2),C5=(3,5)</summary>
    public static string GetCellFromCoordinates(int column, int row)
    {
        return $"{GetAlphabeticalIndex(column - 1)}{row}";
    }

    /**The code below was copied and modified from ExcelDataReader - @MIT License**/
    /// <summary>
    /// Converts Excel cells into numerical coordinates. ex:A2=(1,2),C5=(3,5)
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="column">The column, 1-based.</param>
    /// <param name="row">The row, 1-based.</param>
    public static bool TryParseCellReference(string? value, out int column, out int row)
    {
        const int offset = 'A' - 1;

        row = 0;
        column = 0;
        var position = 0;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        while (position < value!.Length)
        {
            var c = char.ToUpperInvariant(value[position]);
            if (c is >= 'A' and <= 'Z')
            {
                position++;
                column = column * 26 + c - offset;
                continue;
            }

            if (char.IsLetter(c))
            {
                row = 0;
                column = 0;
                position = 0;
                break;
            }

            if (char.IsDigit(c))
                break;

            return false;
        }

        if (position == 0)
            return false;

        return int.TryParse(value[position..], NumberStyles.None, CultureInfo.InvariantCulture, out row) && row > 0;
    }
}
