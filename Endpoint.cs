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
            var lastColumnUsed = worksheet.LastColumnUsed();

            if (lastColumnUsed == null)
                continue;

            var firstRow = worksheet.FirstRowUsed()?.RowNumber() ?? 1;
            var firstCol = worksheet.FirstColumnUsed()?.ColumnNumber() ?? 1;
            var lastCol = lastColumnUsed.ColumnNumber();

            lastCol = worksheet.MergedRanges
                .Select(mergedRange => mergedRange.LastColumn().ColumnNumber())
                .Prepend(lastCol).Max();

            var content = new Dictionary<string, Dictionary<string, object>>();
            for (var col = firstCol; col <= lastCol; col++)
            {
                content[col.ToString()] = new Dictionary<string, object>();
                var lastRowInColumn = firstRow;

                var lastCellUsed = worksheet.Column(col).LastCellUsed();
                if (lastCellUsed != null)
                    lastRowInColumn = lastCellUsed.Address.RowNumber;

                lastRowInColumn = (from mergedRange in worksheet.MergedRanges
                    let firstMergedCol = mergedRange.FirstColumn().ColumnNumber()
                    let lastMergedCol = mergedRange.LastColumn().ColumnNumber()
                    where col >= firstMergedCol && col <= lastMergedCol
                    select mergedRange.LastRow().RowNumber()).Prepend(lastRowInColumn).Max();

                for (var row = firstRow; row <= lastRowInColumn; row++)
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
            return "null";

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