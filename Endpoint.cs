using ClosedXML.Excel;

namespace Converter;

public static class Endpoint
{
    public static void MapEndpoint(this WebApplication app)
        => app.MapPost("export-json", HandlerAsync)
            .Accepts<IFormFile>("multipart/form-data")
            .DisableAntiforgery();

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
            if (!worksheet.RowsUsed().Any())
                continue;

            var content = new Dictionary<string, Dictionary<string, object>>();
            var usedRange = worksheet.RangeUsed();

            if (usedRange == null)
                continue;

            var firstRow = usedRange.FirstRow().RowNumber();
            var lastRow = usedRange.LastRow().RowNumber();
            var firstCol = usedRange.FirstColumn().ColumnNumber();
            var lastCol = usedRange.LastColumn().ColumnNumber();

            for (var row = firstRow; row <= lastRow; row++)
            {
                var rowData = new Dictionary<string, object>();

                for (var col = firstCol; col <= lastCol; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    object cellValue;

                    if (cell.IsMerged())
                    {
                        var mergedRange = cell.MergedRange();
                        cellValue = GetCellValue(mergedRange.FirstCell());
                    }
                    else
                    {
                        cellValue = GetCellValue(cell);
                    }

                    if (cell.IsMerged() || !cell.IsEmpty())
                        rowData[col.ToString()] = cellValue;
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

    private static void Validate(XLWorkbook workbook)
    {
        var ws = workbook.Worksheet(1);
        var rows = ws.RowsUsed();

        if (!rows.Any())
            throw new BadHttpRequestException("Nenhum arquivo enviado");
    }
    
    private static object GetCellValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return null;

        if (cell.Value.IsText)
            return cell.Value.GetText();
    
        if (cell.Value.IsNumber)
            return cell.Value.GetNumber();
    
        if (cell.Value.IsBoolean)
            return cell.Value.GetBoolean();
    
        if (cell.Value.IsDateTime)
            return cell.Value.GetDateTime();
    
        if (cell.Value.IsTimeSpan)
            return cell.Value.GetTimeSpan();
    
        if (cell.Value.IsError)
            return cell.Value.GetError();

        return cell.GetString();
    }
}