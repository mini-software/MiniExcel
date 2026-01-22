namespace MiniExcelLib.OpenXml.Tests.Utils;

internal static class EpplusLicence
{
    static EpplusLicence()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
}