namespace MiniExcelLib.OpenXml.Utils;

internal static class GeneralHelper
{
    public static int GetCellColumnIndex(string cell)
    {
        const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const int mode = 26;

        var x = 0;
        var cellLetter = GetCellColumnLetter(cell);
        //AA=27,ZZ=702
        foreach (var t in cellLetter)
            x = x * mode + keys.IndexOf(t);

        return x;
    }

    public static int GetCellRowNumber(string cell)
    {
        if (string.IsNullOrEmpty(cell))
            throw new Exception("cell is null or empty");
        
        var cellNumber = string.Empty;
        foreach (var t in cell)
        {
            if (char.IsDigit(t))
                cellNumber += t;
        }
        return int.Parse(cellNumber);
    }

    public static string GetCellColumnLetter(string cell)
    {
        string GetCellLetter = string.Empty;
        foreach (var t in cell)
        {
            if (char.IsLetter(t))
                GetCellLetter += t;
        }
        return GetCellLetter;
    }

    public static string ConvertColumnName(int x)
    {
        int dividend = x;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }
}