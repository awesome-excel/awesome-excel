using AwesomeExcel.Models;

namespace AwesomeExcel.Core.Services;

/// <summary>
/// Defines a factory of objects of type Workbook.
/// </summary>
internal class WorkbookFactory
{
    /// <summary>
    /// Create a new Workbook object.
    /// </summary>
    /// <param name="sheets">The sheets of the Workbook.</param>
    /// <param name="customization">Additional customizations used to customize the workbook.</param>
    /// <returns>A new workbook containing the given sheets with the given customizations.</returns>
    /// <exception cref="ArgumentNullException">sheets is null</exception>
    /// /// <exception cref="InvalidOperationException">one or more elements of sheets are null</exception>
    public Workbook Create(IReadOnlyList<Sheet> sheets, WorkbookCustomization? customization)
    {
        if (sheets is null)
        {
            throw new ArgumentNullException(nameof(sheets));
        }

        if (sheets.Count == 0)
        {
            throw new InvalidOperationException(nameof(sheets));
        }

        if (sheets.Any(s => s is null))
        {
            throw new InvalidOperationException(nameof(sheets));
        }

        return new Workbook
        {
            FileType = customization?.FileType ?? FileType.Xlsx,
            Sheets = sheets.ToList()
        };
    }
}
