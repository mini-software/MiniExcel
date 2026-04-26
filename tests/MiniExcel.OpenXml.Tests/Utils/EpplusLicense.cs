namespace MiniExcelLib.OpenXml.Tests.Utils;

internal static class EpplusLicence
{
    internal static void SetContext() 
        => ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
}