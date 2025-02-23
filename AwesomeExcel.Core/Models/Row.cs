﻿namespace AwesomeExcel.Models;

/// <summary>
/// Represents a row of a sheet.
/// </summary>
public class Row
{
    /// <summary>
    /// Cells of the row.
    /// </summary>
    public IEnumerable<Cell> Cells { get; set; }

    /// <summary>
    /// Gets or sets the style of the row.
    /// </summary>
    public Style Style { get; set; }
}
