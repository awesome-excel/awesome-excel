namespace AwesomeExcel.Models;

/// <summary>
/// Represents a style for a font.
/// </summary>
public class FontStyle
{
    /// <summary>
    /// Gets or sets the name of the font.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the color of the font.
    /// </summary>
    public Color? Color { get; set; }

    /// <summary>
    /// Gets or sets the height in points of the font.
    /// </summary>
    public short? HeightInPoints { get; set; }

    /// <summary>
    /// Determines whether the font is bold.
    /// <br /> Null makes this style to be inherited from parent.
    /// </summary>
    public bool? IsBold { get; set; }

    public FontStyle ShallowCopy()
    {
        return (FontStyle)MemberwiseClone();
    }
}
