namespace AwesomeExcel.Models;

/// <summary>
/// Represents a column of a sheet.
/// </summary>
public class Column
{
    /// <summary>
    /// Gets or sets the name of the column.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the type of the column.
    /// </summary>
    public ColumnType ColumnType { get; set; }

    /// <summary>
    /// Gets or sets the style of the column.
    /// </summary>
    public Style Style { get; set; }
}
