using OfficeOpenXml;

namespace MiniExcelLibs.Tests.Utils;

internal static class EpplusLicence
{
    static EpplusLicence()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }
}