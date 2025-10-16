using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace Converter;

public static class EndpointV2
{
    public static void MapEndpointV2(this WebApplication app)
        => app.MapPost("export-json-v2", HandlerAsync)
            .Accepts<IFormFile>("multipart/form-data")
            .DisableAntiforgery();

    private static async Task<IResult> HandlerAsync(
        IFormFile file,
        [FromForm] string? password,
        CancellationToken cancellationToken)
    {
        if (file.Length == 0)
            return Results.BadRequest("Nenhum arquivo enviado");

        ExcelPackage.License.SetNonCommercialPersonal("Lucas de Lima Canto");
        await using var stream = file.OpenReadStream();

        try
        {
            using var package = string.IsNullOrEmpty(password)
                ? new ExcelPackage(stream)
                : new ExcelPackage(stream, password);
            // using var package = new ExcelPackage(stream, "adasda");
            var sheetsData = new List<object>();

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var dimension = worksheet.Dimension;

                if (dimension == null)
                    continue;

                var firstRow = dimension.Start.Row;
                var firstCol = dimension.Start.Column;
                var lastCol = dimension.End.Column;
                var content = new Dictionary<string, Dictionary<string, object>>();

                for (var col = firstCol; col <= lastCol; col++)
                {
                    content[col.ToString()] = new Dictionary<string, object>();
                    var lastRowInColumn = firstRow;

                    for (var row = dimension.End.Row; row >= firstRow; row--)
                    {
                        var cell = worksheet.Cells[row, col];

                        if (!cell.Merge && cell.Value == null) continue;
                        lastRowInColumn = row;
                        break;
                    }

                    lastRowInColumn = (from mergeAddress in worksheet.MergedCells
                        select worksheet.Cells[mergeAddress]
                        into mergeRange
                        let firstMergedCol = mergeRange.Start.Column
                        let lastMergedCol = mergeRange.End.Column
                        where col >= firstMergedCol && col <= lastMergedCol
                        select mergeRange.End.Row).Prepend(lastRowInColumn).Max();

                    for (var row = firstRow; row <= lastRowInColumn; row++)
                    {
                        var cell = worksheet.Cells[row, col];
                        object? cellValue = null;

                        if (cell.Merge)
                        {
                            var mergeId = worksheet.GetMergeCellId(row, col);
                            if (mergeId > 0)
                            {
                                var mergeAddress = worksheet.MergedCells[mergeId - 1];
                                var mergeRange = worksheet.Cells[mergeAddress];
                                var firstCell = worksheet.Cells[mergeRange.Start.Row, mergeRange.Start.Column];
                                cellValue = GetCellValue(firstCell);
                            }
                        }
                        else
                            cellValue = GetCellValue(cell);

                        content[col.ToString()][row.ToString()] = cellValue;
                    }
                }

                sheetsData.Add(new
                {
                    sheet = worksheet.Name,
                    content
                });
            }

            return Results.Ok(new { sheets = sheetsData });
        }
        catch (InvalidDataException)
        {
            return Results.BadRequest("Arquivo inválido ou senha incorreta");
        }
        catch (Exception ex)
        {
            return Results.Problem($"Erro ao processar arquivo: {ex.Message}");
        }
    }

    private static object GetCellValue(ExcelRange cell)
    {
        var value = cell.Value;

        if (value == null)
            return "null";

        return value switch
        {
            DateTime dt => dt,
            double d => d,
            int i => i,
            bool b => b,
            _ => value.ToString()
        };
    }
}