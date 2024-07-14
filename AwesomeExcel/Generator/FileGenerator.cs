using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Services;

namespace AwesomeExcel.Generator;

/// <summary>
/// A base class for an Excel file generator.
/// </summary>
/// <typeparam name="TWorkbook"></typeparam>
public abstract class FileGenerator<TWorkbook>
{
    private readonly WorkbookFactory workbookFactory = new();
    private readonly SheetFactory sheetFactory = new();

    /// <summary>
    /// Generate an Excel File.
    /// </summary>
    /// <typeparam name="TSheet">The type of the worksheet rows.</typeparam>
    /// <param name="rows">The rows of the sheet.</param>
    /// <param name="customization">A delegate used to customize the Excel file.</param>
    /// <returns>The MemoryStream of the Excel file.</returns>
    public MemoryStream Generate<TSheet>(IEnumerable<TSheet> rows, Action<SheetCustomizer<TSheet>> customization)
    {
        SheetCustomizer<TSheet> customizer = new();
        customization(customizer);

        Sheet sheet = sheetFactory.Create(rows, customizer.Sheet, customizer.GetCustomizedColumns(), customizer.GetCustomizedCells());

        Workbook excelWorkbook = workbookFactory.Create(new Sheet[1] { sheet }, customizer.Workbook);
        return GetStream(excelWorkbook);
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
    public MemoryStream Generate<TSheet1, TSheet2>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, Action<SheetsCustomizer<TSheet1, TSheet2>> customization)
    {
        SheetsCustomizer<TSheet1, TSheet2> customizer = new();
        customization(customizer);

        Sheet sheet1 = sheetFactory.Create(rowsSheet1, customizer.Sheet1, customizer.GetCustomizedColumns(customizer.Sheet1), customizer.GetCustomizedCells(customizer.Sheet1));
        Sheet sheet2 = sheetFactory.Create(rowsSheet2, customizer.Sheet2, customizer.GetCustomizedColumns(customizer.Sheet2), customizer.GetCustomizedCells(customizer.Sheet2));

        Workbook excelWorkbook = workbookFactory.Create(new Sheet[] { sheet1, sheet2 }, customizer.Workbook);
        return GetStream(excelWorkbook);
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
    public MemoryStream Generate<TSheet1, TSheet2, TSheet3>(IEnumerable<TSheet1> rowsSheet1, IEnumerable<TSheet2> rowsSheet2, IEnumerable<TSheet3> rowsSheet3, Action<SheetsCustomizer<TSheet1, TSheet2, TSheet3>> customization)
    {
        SheetsCustomizer<TSheet1, TSheet2, TSheet3> customizer = new();
        customization(customizer);

        Sheet sheet1 = sheetFactory.Create(rowsSheet1, customizer.Sheet1, customizer.GetCustomizedColumns(customizer.Sheet1), customizer.GetCustomizedCells(customizer.Sheet1));
        Sheet sheet2 = sheetFactory.Create(rowsSheet2, customizer.Sheet2, customizer.GetCustomizedColumns(customizer.Sheet2), customizer.GetCustomizedCells(customizer.Sheet2));
        Sheet sheet3 = sheetFactory.Create(rowsSheet3, customizer.Sheet3, customizer.GetCustomizedColumns(customizer.Sheet3), customizer.GetCustomizedCells(customizer.Sheet3));

        Workbook excelWorkbook = workbookFactory.Create(new List<Sheet> { sheet1, sheet2, sheet3 }, customizer.Workbook);
        return GetStream(excelWorkbook);
    }

    private MemoryStream GetStream(Workbook workbook)
    {
        TWorkbook serviceWorkbook = Convert(workbook);
        MemoryStream stream = Write(serviceWorkbook);
        return stream;
    }

    /// <summary>
    /// Convert a common Workbook to a library specific implementation of a Workbook.
    /// </summary>
    /// <param name="workbook">The workbook to be converted.</param>
    /// <returns>The converted workbook.</returns>
    protected abstract TWorkbook Convert(Workbook workbook);

    /// <summary>
    /// Write a specific implementation of a Workbook to a stream.
    /// </summary>
    /// <param name="workbook">The workbook to be written.</param>
    /// <returns>The stream the Excel file is written into.</returns>
    protected abstract MemoryStream Write(TWorkbook workbook);
}


