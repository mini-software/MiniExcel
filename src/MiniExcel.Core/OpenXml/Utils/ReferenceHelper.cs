namespace MiniExcelLib.Core.OpenXml.Utils;

internal static class ReferenceHelper
{
    public static string GetCellNumber(string cell)
    {
        return cell
            .Where(char.IsDigit)
            .Aggregate(string.Empty, (current, c) => current + c);
    }

    public static string GetCellLetter(string cell)
    {
        return cell
            .Where(char.IsLetter)
            .Aggregate(string.Empty, (current, c) => current + c);
    }

    /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
    public static (int, int) ConvertCellToCoordinates(string? cell)
    {
        const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const int mode = 26;

        //AA=27,ZZ=702
        var cellLetter = GetCellLetter(cell);
        var x = cellLetter.Aggregate(0, (idx, chr) => idx * mode + keys.IndexOf(chr));

        var cellNumber = GetCellNumber(cell);
        return (x, int.Parse(cellNumber));
    }

    /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
    public static string ConvertCoordinatesToCell(int x, int y)
    {
        int dividend = x;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return $"{columnName}{y}";
    }

    /// <summary>
    /// Try to parse cell reference (e.g., "A1") into column and row numbers.
    /// </summary>
    /// <param name="cellRef">The cell reference (e.g., "A1", "B2", "AA10")</param>
    /// <param name="column">The column number (1-based)</param>
    /// <param name="row">The row number (1-based)</param>
    /// <returns>True if successfully parsed, false otherwise</returns>
    public static bool TryParseCellReference(string cellRef, out int column, out int row)
    {
        column = 0;
        row = 0;
        
        if (string.IsNullOrEmpty(cellRef))
            return false;
        
        try
        {
            var coords = ConvertCellToCoordinates(cellRef);
            column = coords.Item1;
            row = coords.Item2;
            return column > 0 && row > 0;
        }
        catch
        {
            return false;
        }
    }
    
    /**The code below was copied and modified from ExcelDataReader - @MIT License**/
    /// <summary>
    /// Logic for the Excel dimensions. Ex: A15
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="column">The column, 1-based.</param>
    /// <param name="row">The row, 1-based.</param>
    public static bool ParseReference(string value, out int column, out int row)
    {
        const int offset = 'A' - 1;
        
        row = 0;
        column = 0;
        var position = 0;

        if (string.IsNullOrWhiteSpace(value))
            return false;

        while (position < value.Length)
        {
            var c = char.ToUpperInvariant(value[position]);
            if (c is >= 'A' and <= 'Z')
            {
                position++;
                column *= 26;
                column += c - offset;
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