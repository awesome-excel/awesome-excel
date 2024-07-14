using AwesomeExcel.BridgeNpoi;
using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Services;
using AwesomeExcel.Generator;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace AwesomeExcel;

public class AwesomeExcel
{
    private readonly NpoiFileGenerator generator = new();

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet">The type of the worksheet rows.</typeparam>
    /// <param name="rows">The rows of the sheet.</param>
    /// <param name="action">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet>(IEnumerable<TSheet> rows, Action<SheetCustomizer<TSheet>> action)
    {
        return generator.Generate(rows, action);
    }

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet1">The type of the rows of the first sheet.</typeparam>
    /// <typeparam name="TSheet2">The type of the rows of the second sheet.</typeparam>
    /// <param name="rowsSheet1">The rows of the first sheet.</param>
    /// <param name="rowsSheet2">The rows of the second sheet.</param>
    /// <param name="action">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet1, TSheet2>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, Action<SheetsCustomizer<TSheet1, TSheet2>> action)
    {
        return generator.Generate(rowsSheet1, rowsSheet2, action);
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
    /// <param name="action">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet1, TSheet2, TSheet3>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, IEnumerable<TSheet3> rowsSheet3, Action<SheetsCustomizer<TSheet1, TSheet2, TSheet3>> action)
    {
        return generator.Generate(rowsSheet1, rowsSheet2, rowsSheet3, action);
    }

}
