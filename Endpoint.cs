using ClosedXML.Excel;

namespace Converter;

public static class Endpoint
{
    public static void MapEndpoint(this WebApplication app)
        => app.MapPost("export-json", HandlerAsync)
            .WithName("export-json")
            .WithDisplayName("transformar excel em json na teoria");

    private static async Task<IResult> HandlerAsync(
        IFormFile file,
        CancellationToken cancellationToken)
    {
        if (file.Length == 0)
            return Results.BadRequest("Nenhum arquivo enviado");

        await using var stream = file.OpenReadStream();
        using var workbook = new XLWorkbook(stream);
        Validate(workbook);
        var sheetsData = new List<object>();

        foreach (var worksheet in workbook.Worksheets)
        {
            var content = new Dictionary<string, Dictionary<string, object>>();
            if (!worksheet.RowsUsed().Any())
                continue;

            var columnsCount = worksheet.ColumnsUsed().Count();
            var rowsCount = worksheet.LastRow().RowNumber();

            for (var row = 1; row <= rowsCount; row++)
            {
                var rowData = new Dictionary<string, object>();
                for (var col = 1; col <= columnsCount; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    rowData[col.ToString()] = cell.Value;
                }

                content[row.ToString()] = rowData;
            }

            sheetsData.Add(new
            {
                sheet = worksheet.Name,
                content
            });
        }

        return Results.Ok(new { sheets = sheetsData });
    }

    private static void Validate(XLWorkbook workbook)
    {
        var ws = workbook.Worksheet(1);
        var rows = ws.RowsUsed();

        if (!rows.Any())
            throw new BadHttpRequestException("Nenhum arquivo enviado");
    }
}