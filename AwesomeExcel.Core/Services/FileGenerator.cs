namespace AwesomeExcel.Core.Services;

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
    public MemoryStream Generate<TSheet>(IEnumerable<TSheet> rows, Action<SheetCustomizer<TSheet>> customization = null)
    {
        SheetCustomizer<TSheet> customizer = GetCustomizer(customization);

        Sheet sheet = sheetFactory.Create(rows, customizer, customizer?.GetColumns(), customizer?.GetCells());
        Workbook workbook = workbookFactory.Create(new Sheet[1] { sheet }, customizer?.Workbook);
        return GetStream(workbook);
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
        var (customizer1, customizer2) = GetCustomizer(customization);

        Sheet sheet1 = sheetFactory.Create(rowsSheet1, customizer1, customizer1?.GetColumns(), customizer1?.GetCells());
        Sheet sheet2 = sheetFactory.Create(rowsSheet2, customizer2, customizer2?.GetColumns(), customizer2?.GetCells());

        Workbook workbook = workbookFactory.Create(new Sheet[] { sheet1, sheet2 }, customizer1?.Workbook);
        return GetStream(workbook);
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
        var (customizer1, customizer2, customizer3) = GetCustomizer(customization);

        Sheet sheet1 = sheetFactory.Create(rowsSheet1, customizer1, customizer1?.GetColumns(), customizer1?.GetCells());
        Sheet sheet2 = sheetFactory.Create(rowsSheet2, customizer2, customizer2?.GetColumns(), customizer2?.GetCells());
        Sheet sheet3 = sheetFactory.Create(rowsSheet3, customizer3, customizer3?.GetColumns(), customizer3?.GetCells());

        Workbook workbook = workbookFactory.Create(new List<Sheet> { sheet1, sheet2, sheet3 }, customizer1?.Workbook);
        return GetStream(workbook);
    }

    private MemoryStream GetStream(Workbook workbook)
    {
        TWorkbook serviceWorkbook = Convert(workbook);
        MemoryStream stream = Write(serviceWorkbook);

        if (typeof(TWorkbook).IsAssignableTo(typeof(IDisposable)))
        {
            ((IDisposable)serviceWorkbook)?.Dispose();
        }

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

    private SheetCustomizer<TSheet> GetCustomizer<TSheet>(Action<SheetCustomizer<TSheet>> customization)
    {
        if (customization == null)
        {
            return null;
        }

        SheetCustomizer<TSheet> customizer = new();
        customization(customizer);
        return customizer;
    }

    private (SheetCustomizer<TSheet1>, SheetCustomizer<TSheet2>) GetCustomizer<TSheet1, TSheet2>(Action<SheetCustomizer<TSheet1>, SheetCustomizer<TSheet2>> customization)
    {
        if (customization == null)
        {
            return (null, null);
        }

        SheetCustomizer<TSheet1> customizer1 = new();
        SheetCustomizer<TSheet2> customizer2 = new();

        customization(customizer1, customizer2);
        return (customizer1, customizer2);
    }

    private (SheetCustomizer<TSheet1>, SheetCustomizer<TSheet2>, SheetCustomizer<TSheet3>) GetCustomizer<TSheet1, TSheet2, TSheet3>(Action<SheetCustomizer<TSheet1>, SheetCustomizer<TSheet2>, SheetCustomizer<TSheet3>> customization)
    {
        if (customization == null)
        {
            return (null, null, null);
        }

        SheetCustomizer<TSheet1> customizer1 = new();
        SheetCustomizer<TSheet2> customizer2 = new();
        SheetCustomizer<TSheet3> customizer3 = new();

        customization(customizer1, customizer2, customizer3);
        return (customizer1, customizer2, customizer3);
    }
}


