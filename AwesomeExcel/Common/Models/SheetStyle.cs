namespace AwesomeExcel.Common.Models;

/// <summary>
/// Represents a style for sheet.
/// </summary>
public class SheetStyle : Style
{
    /// <summary>
    /// Gets or sets the colors used for even and odd rows.
    /// </summary>
    public ColorBanding ColorBanding { get; set; }
}
