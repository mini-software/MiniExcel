namespace MiniExcelLib.Core.Helpers;

internal static class StringHelper
{
    public static string GetLetters(string content) => new([..content.Where(char.IsLetter)]);
    public static int GetNumber(string content) => int.Parse(new string([..content.Where(char.IsNumber)]));
}