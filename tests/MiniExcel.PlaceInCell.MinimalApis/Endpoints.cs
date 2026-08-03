using MiniExcelLib.Core;
using MiniExcelLib.Core.Enums;
using MiniExcelLib.OpenXml.Api;
using MiniExcelLib.OpenXml.Picture;

namespace MiniExcel.PlaceInCell.MinimalApis;

internal static class Endpoints
{
    private static readonly OpenXmlExporter Exporter = MiniExcelLib.Core.MiniExcel.Exporters.GetOpenXmlExporter();
    private static readonly OpenXmlTemplater Templater = MiniExcelLib.Core.MiniExcel.Templaters.GetOpenXmlTemplater();

    private const string ExcelContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    internal static RouteGroupBuilder MapPlaceInCellApi(this IEndpointRouteBuilder builder)
    {
        var group = builder.MapGroup("api/place-in-cell");

        group.MapStreamExport();
        group.MapSaveToDisk();

        return group;
    }

    /// <summary>
    /// Generate xlsx in memory (no disk) and return as download stream.
    /// </summary>
    private static RouteHandlerBuilder MapStreamExport(this IEndpointRouteBuilder builder)
    {
        return builder.MapGet("stream", async (CancellationToken cancellationToken) =>
        {
            var memoryStream = new MemoryStream();
            await CreatePlaceInCellWorkbookAsync(memoryStream, cancellationToken);

            memoryStream.Seek(0, SeekOrigin.Begin);
            return Results.Stream(memoryStream, ExcelContentType, "place-in-cell-stream.xlsx");
        })
        .WithName("PlaceInCellStream")
        .WithSummary("Generate Place in Cell xlsx in memory and download")
        .WithDescription("Exports a workbook with PlaceInCell images using MemoryStream only (no file on disk).")
        .Produces(StatusCodes.Status200OK, contentType: ExcelContentType);
    }

    /// <summary>
    /// Save xlsx to the desktop and return the file path.
    /// </summary>
    private static RouteHandlerBuilder MapSaveToDisk(this IEndpointRouteBuilder builder)
    {
        return builder.MapGet("save", async (CancellationToken cancellationToken) =>
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var path = Path.Combine(desktop, $"place-in-cell-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");

            // Prefer path-based APIs (same pattern as unit tests / OpenXmlTemplater file overloads)
            var rows = CreateDemoRows();
            await Exporter.ExportAsync(path, rows, cancellationToken: cancellationToken);

            var images = await CreateDemoPicturesAsync(cancellationToken);
            await Templater.AddPictureAsync(path, cancellationToken, images);

            return Results.Ok(new
            {
                message = "Saved to desktop. Open with Microsoft 365 Excel to verify Place in Cell.",
                path
            });
        })
        .WithName("PlaceInCellSave")
        .WithSummary("Save Place in Cell xlsx to desktop")
        .WithDescription("Writes the workbook to the desktop via file-path Export + AddPicture, then returns the saved path.")
        .Produces(StatusCodes.Status200OK);
    }

    private static async Task CreatePlaceInCellWorkbookAsync(Stream stream, CancellationToken cancellationToken)
    {
        var rows = CreateDemoRows();
        await Exporter.ExportAsync(stream, rows, cancellationToken: cancellationToken);
        stream.Seek(0, SeekOrigin.Begin);

        var images = await CreateDemoPicturesAsync(cancellationToken);
        await Templater.AddPictureAsync(stream, cancellationToken, images);
        stream.Seek(0, SeekOrigin.Begin);
    }

    private static object[] CreateDemoRows() =>
    [
        new { Product = "GitHub", Note = "logo in C2" },
        new { Product = "Google", Note = "logo in C3" }
    ];

    private static async Task<MiniExcelPicture[]> CreateDemoPicturesAsync(CancellationToken cancellationToken)
    {
        var github = await File.ReadAllBytesAsync(
            Path.Combine(AppContext.BaseDirectory, "images", "github_logo.png"), cancellationToken);
        var google = await File.ReadAllBytesAsync(
            Path.Combine(AppContext.BaseDirectory, "images", "google_logo.png"), cancellationToken);

        return
        [
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
            }
        ];
    }
}
