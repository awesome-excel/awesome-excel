using AwesomeExcel.BridgeNPOI;

namespace AwesomeExcel;

public class ExcelGenerator 
{
    private NpoiFileGenerator fileGenerator = new();

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet">The type of the worksheet rows.</typeparam>
    /// <param name="rows">The rows of the sheet.</param>
    /// <param name="customization">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet>(IEnumerable<TSheet> rows, Action<SheetCustomizer<TSheet>> customization = null)
    {
        return fileGenerator.Generate(rows, customization);
    }

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet1">The type of the rows of the first sheet.</typeparam>
    /// <typeparam name="TSheet2">The type of the rows of the second sheet.</typeparam>
    /// <param name="rowsSheet1">The rows of the first sheet.</param>
    /// <param name="rowsSheet2">The rows of the second sheet.</param>
    /// <param name="customization">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet1, TSheet2>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, Action<SheetCustomizer<TSheet1>, SheetCustomizer<TSheet2>> customization = null)
    {
        return fileGenerator.Generate(rowsSheet1, rowsSheet2, customization);
    }

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet1">The type of the rows of the first sheet.</typeparam>
    /// <typeparam name="TSheet2">The type of the rows of the second sheet.</typeparam>
    /// <typeparam name="TSheet3">The type of the rows of the third sheet.</typeparam>
    /// <param name="rowsSheet1">The rows of the first sheet.</param>
    /// <param name="rowsSheet2">The rows of the second sheet.</param>
    /// <param name="rowsSheet3">The rows of the third sheet.</param>
    /// <param name="customization">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet1, TSheet2, TSheet3>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, IEnumerable<TSheet3> rowsSheet3, Action<SheetCustomizer<TSheet1>, SheetCustomizer<TSheet2>, SheetCustomizer<TSheet3>> customization)
    {
        return fileGenerator.Generate(rowsSheet1, rowsSheet2, rowsSheet3, customization);
    }

}
