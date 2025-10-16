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

            var sheetsData = new List<object>();

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var dimension = worksheet.Dimension;
                
                if (dimension == null)
                    continue;

                var content = new Dictionary<string, Dictionary<string, object>>();
                var firstRow = dimension.Start.Row;
                var lastRow = dimension.End.Row;
                var firstCol = dimension.Start.Column;
                var lastCol = dimension.End.Column;

                for (var row = firstRow; row <= lastRow; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    for (var col = firstCol; col <= lastCol; col++)
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


                        if (cell.Merge || cellValue != null)
                            rowData[col.ToString()] = cellValue ?? string.Empty;
                    }

                    if (rowData.Count != 0)
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
        catch (InvalidDataException)
        {
            return Results.BadRequest("Arquivo inválido ou senha incorreta");
        }
        catch (Exception ex)
        {
            return Results.Problem($"Erro ao processar arquivo: {ex.Message}");
        }
    }

    private static object? GetCellValue(ExcelRange cell)
    {
        var value = cell.Value;
        
        if (value == null)
            return null;

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