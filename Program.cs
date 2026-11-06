using ClosedXML.Excel;
using System.Globalization;
using System.Text;
string fich ="";

Console.OutputEncoding = Encoding.UTF8;
if(args.Length != 1)
{
    Console.Write("Fichero Original (.xlsx): ");
    fich = Console.ReadLine()?.Trim() ?? "";
}
else
{
    fich = args[0];
}
string? temporaryPath = null;
try
{
    string inputPath = Path.GetFullPath(fich.Trim('"'));
    if (!File.Exists(inputPath))
        throw new FileNotFoundException("No existe el archivo indicado.", inputPath);
    if (!Path.GetExtension(inputPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        throw new ArgumentException("El archivo debe tener extensión .xlsx.");

    Console.Write("A partir de la fila: ");
    if (!int.TryParse(Console.ReadLine(), out int startRow) || startRow < 1 || startRow > 1_048_576)
        throw new ArgumentException("La fila inicial debe ser un entero entre 1 y 1048576.");
    Console.Write("Número de contenedor: ");
    string container = Console.ReadLine()?.Trim() ?? "";
    if (string.IsNullOrWhiteSpace(container))
        throw new ArgumentException("El número de contenedor es obligatorio.");

    string outputPath = Path.Combine(Path.GetDirectoryName(inputPath)!,
        Path.GetFileNameWithoutExtension(inputPath) + "_convertido.xlsx");
    if (File.Exists(outputPath))
        throw new IOException($"Ya existe {outputPath}. Muévalo o elimínelo antes de repetir la conversión.");

    using var source = new XLWorkbook(inputPath);
    var sheet = source.Worksheet(source.Worksheets.Count);
    int lastRow = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
    if (startRow > lastRow)
        throw new ArgumentException("La fila inicial está después de la última fila con datos.");

    using var result = new XLWorkbook();
    var detail = result.AddWorksheet("Detalle");
    var summary = result.AddWorksheet("Resumen");
    SetHeaders(detail, "Codigo", "Producto", "Cantidad", "Lote");
    SetHeaders(summary, "Codigo", "Producto", "Lote", "Total Cantidad");
    detail.Column(1).Style.NumberFormat.Format = "@";
    detail.Column(4).Style.NumberFormat.Format = "@";
    summary.Column(1).Style.NumberFormat.Format = "@";
    summary.Column(3).Style.NumberFormat.Format = "@";

    var totals = new Dictionary<(string Code, string Product, string Lot), decimal>();
    int outputRow = 2;
    for (int row = startRow; row <= lastRow; row++)
    {
        string code = ReadText(sheet.Cell(row, 2));
        string product = ReadText(sheet.Cell(row, 3));
        for (int offset = 0; offset < 6; offset++)
        {
            var cell = sheet.Cell(row, 8 + offset);
            if (cell.DataType == XLDataType.Blank) continue;
            decimal quantity = ReadQuantity(cell);
            string lot = ReadText(sheet.Cell(row, 14 + offset));
            if (lot.Length > 0) lot = $"{container}-{lot}";
            if (outputRow > 1_048_576)
                throw new InvalidOperationException("El detalle supera el límite de filas de Excel.");
            detail.Cell(outputRow, 1).Value = code;
            detail.Cell(outputRow, 2).Value = product;
            detail.Cell(outputRow, 3).Value = quantity;
            detail.Cell(outputRow, 4).Value = lot;
            outputRow++;
            var key = (code, product, lot);
            totals[key] = checked(totals.GetValueOrDefault(key) + quantity);
        }
    }
    if (totals.Count == 0)
        throw new InvalidOperationException("No se encontraron cantidades en H:M desde la fila indicada.");

    int summaryRow = 2;
    foreach (var entry in totals)
    {
        summary.Cell(summaryRow, 1).Value = entry.Key.Code;
        summary.Cell(summaryRow, 2).Value = entry.Key.Product;
        summary.Cell(summaryRow, 3).Value = entry.Key.Lot;
        summary.Cell(summaryRow, 4).Value = entry.Value;
        summaryRow++;
    }
    FormatSheet(detail, outputRow - 1, 3);
    FormatSheet(summary, summaryRow - 1, 4);

    // Guardar primero en un temporal evita dejar un resultado incompleto.
    temporaryPath = Path.Combine(Path.GetDirectoryName(outputPath)!, $".{Guid.NewGuid():N}.xlsx");
    result.SaveAs(temporaryPath);
    File.Move(temporaryPath, outputPath, overwrite: false);
    temporaryPath = null;
    Console.WriteLine($"Archivo creado: {outputPath}");
    Console.WriteLine($"Detalle: {outputRow - 2} registros. Resumen: {totals.Count} grupos.");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}
finally
{
    if (temporaryPath is not null)
    {
        try { File.Delete(temporaryPath); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}

static string ReadText(IXLCell cell) => cell.GetFormattedString(CultureInfo.InvariantCulture).Trim();

static decimal ReadQuantity(IXLCell cell)
{
    if (cell.DataType == XLDataType.Number && cell.TryGetValue<decimal>(out var numeric))
        return numeric;
    // Texto: coma o punto decimal, sin separadores de miles para evitar ambigüedad.
    if (cell.DataType == XLDataType.Text)
    {
        string text = cell.GetString().Trim().Replace(',', '.');
        if (decimal.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out var value)) return value;
    }
    throw new InvalidDataException($"Cantidad no válida en {cell.Address}: '{cell.GetString()}'. Use un número o texto decimal sin separadores de miles.");
}

static void SetHeaders(IXLWorksheet sheet, params string[] headers)
{
    for (int i = 0; i < headers.Length; i++) sheet.Cell(1, i + 1).Value = headers[i];
}

static void FormatSheet(IXLWorksheet sheet, int lastRow, int quantityColumn)
{
    sheet.Range(1, 1, 1, 4).Style.Font.Bold = true;
    sheet.Range(1, 1, 1, 4).Style.Fill.BackgroundColor = XLColor.LightBlue;
    sheet.Range(1, 1, lastRow, 4).SetAutoFilter();
    sheet.SheetView.FreezeRows(1);
    sheet.Column(quantityColumn).Style.NumberFormat.Format = "0.###############";
    sheet.Column(1).Width = 22;
    sheet.Column(2).Width = 45;
    sheet.Column(3).Width = 22;
    sheet.Column(4).Width = 22;
}
