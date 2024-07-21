namespace AwesomeExcel;

/// <summary>
/// Represents a cell of a sheet.
/// </summary>
public class Cell
{
    /// <summary>
    /// Gets or sets the value of the cell.
    /// </summary>
    public object Value { get; set; }

    /// <summary>
    /// Gets or sets the style of the cell.
    /// </summary>
    public Style Style { get; set; }
}

public class Cell<T> : Cell { }