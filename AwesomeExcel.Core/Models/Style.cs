namespace AwesomeExcel.Models;

/// <summary>
/// Represents a generic style.
/// </summary>
public class Style
{
    /// <summary>
    /// Gets or sets the color of the top border.
    /// </summary>
    public Color? BorderTopColor { get; set; }

    /// <summary>
    /// Gets or sets the color of the bottom border.
    /// </summary>
    public Color? BorderBottomColor { get; set; }

    /// <summary>
    /// Gets or sets the color of the left border.
    /// </summary>
    public Color? BorderLeftColor { get; set; }

    /// <summary>
    /// Gets or sets the color of the right border.
    /// </summary>
    public Color? BorderRightColor { get; set; }

    /// <summary>
    /// Gets or sets the color of the foreground.
    /// <br /> FillPattern must be set.
    /// </summary>
    public Color? FillForegroundColor { get; set; }

    /// <summary>
    /// Gets or sets the fill pattern.
    /// </summary>
    public FillPattern? FillPattern { get; set; }

    /// <summary>
    /// Gets or sets the style of the font.
    /// </summary>
    public FontStyle FontStyle { get; set; }

    /// <summary>
    /// Gets or sets the date time format.
    /// </summary>
    public string DateTimeFormat { get; set; }

    /// <summary>
    /// Gets or sets the horizontal alignment.
    /// </summary>
    public HorizontalAlignment? HorizontalAlignment { get; set; }

    /// <summary>
    /// Gets or sets the vertical alignment.
    /// </summary>
    public VerticalAlignment? VerticalAlignment { get; set; }

    public Style ShallowCopy()
    {
        return (Style)MemberwiseClone();
    }
}
