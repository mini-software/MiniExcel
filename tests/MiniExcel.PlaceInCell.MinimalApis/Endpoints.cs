using MiniExcelLib.Core;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.OpenXml.Api;
using MiniExcelLib.OpenXml.Picture;
using CoreMiniExcel = MiniExcelLib.Core.MiniExcel;

namespace MiniExcel.PlaceInCell.MinimalApis;

internal static class Endpoints
{
    private static readonly OpenXmlExporter Exporter = CoreMiniExcel.Exporters.GetOpenXmlExporter();
    private static readonly OpenXmlTemplater Templater = CoreMiniExcel.Templaters.GetOpenXmlTemplater();

    private static readonly string ExcelContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    internal static RouteGroupBuilder MapPlaceInCellApi(this IEndpointRouteBuilder builder)
    {
        var group = builder.MapGroup("api/place-in-cell");

        group.MapGet("stream", GenerateInMemoryAsync);
        group.MapGet("save", SaveToDiskAsync);

        return group;
    }

    /// <summary>
    /// 不落地：内存生成 xlsx（Place in Cell），直接作为文件流下载。
    /// </summary>
    private static async Task<IResult> GenerateInMemoryAsync()
    {
        var stream = await CreatePlaceInCellWorkbookAsync().ConfigureAwait(false);
        stream.Position = 0;
        return Results.File(stream, ExcelContentType, "place-in-cell-stream.xlsx");
    }

    /// <summary>
    /// 落地：保存 xlsx 到桌面，返回保存路径（可用 Excel 365 打开验证嵌入效果）。
    /// </summary>
    private static async Task<IResult> SaveToDiskAsync()
    {
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        var fileName = $"place-in-cell-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx";
        var path = Path.Combine(desktop, fileName);

        await using (var fileStream = File.Create(path))
        {
            await using var workbook = await CreatePlaceInCellWorkbookAsync().ConfigureAwait(false);
            workbook.Position = 0;
            await workbook.CopyToAsync(fileStream).ConfigureAwait(false);
        }

        return Results.Ok(new
        {
            message = "已保存到桌面，请用 Microsoft 365 Excel 打开查看 Place in Cell 效果。",
            path
        });
    }

    private static async Task<MemoryStream> CreatePlaceInCellWorkbookAsync()
    {
        var stream = new MemoryStream();
        var rows = new[]
        {
            new { Product = "GitHub", Note = "logo in C2" },
            new { Product = "Google", Note = "logo in C3" }
        };

        await Exporter.ExportAsync(stream, rows).ConfigureAwait(false);
        stream.Position = 0;

        var github = await File.ReadAllBytesAsync(Path.Combine(AppContext.BaseDirectory, "images", "github_logo.png"))
            .ConfigureAwait(false);
        var google = await File.ReadAllBytesAsync(Path.Combine(AppContext.BaseDirectory, "images", "google_logo.png"))
            .ConfigureAwait(false);

        await Templater.AddPictureAsync(
            stream,
            CancellationToken.None,
            new MiniExcelPicture
            {
                ImageBytes = github,
                CellAddress = "C2",
                ImgType = XlsxImgType.PlaceInCell,
                PictureType = "image/png"
            },
            new MiniExcelPicture
            {
                ImageBytes = google,
                CellAddress = "C3",
                ImgType = XlsxImgType.PlaceInCell,
                PictureType = "image/png"
            }).ConfigureAwait(false);

        stream.Position = 0;
        return stream;
    }
}
