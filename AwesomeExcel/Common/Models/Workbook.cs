namespace AwesomeExcel.Common.Models;

/// <summary>
/// Represents an Excel Workbook.
/// </summary>
public class Workbook
{
    /// <summary>
    /// Gets or sets the sheets of the workbook.
    /// </summary>
    public IList<Sheet> Sheets { get; set; }

    /// <summary>
    /// Gets or sets the file type used for generating the workbook.
    /// </summary>
    public FileType FileType { get; set; }
}
 