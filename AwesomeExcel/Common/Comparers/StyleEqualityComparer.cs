using AwesomeExcel.Common.Models;

namespace AwesomeExcel.Common.Comparers;

/// <summary>
/// Defines methods to support the comparison of objects of type Style for equality.
/// </summary>
public class StyleEqualityComparer : IEqualityComparer<Style>
{
    private readonly FontStyleEqualityComparer fontStyleEqualityComparer;

    public StyleEqualityComparer()
    {
        fontStyleEqualityComparer = new FontStyleEqualityComparer();
    }

    /// <summary>
    /// Determines whether the specified Styles are equal.
    /// </summary>
    /// <param name="x">The first style to compare.</param>
    /// <param name="y">The second style to compare.</param>
    /// <returns>true if the specified styles are equal; otherwise, false.</returns>
    public bool Equals(Style x, Style y)
    {
        if (x == y)
            return true;

        if (x is null || y is null)
            return false;

        return x.BorderTopColor == y.BorderTopColor
            && x.BorderBottomColor == y.BorderBottomColor
            && x.BorderLeftColor == y.BorderLeftColor
            && x.BorderRightColor == y.BorderRightColor
            && x.FillForegroundColor == y.FillForegroundColor
            && x.FillPattern == y.FillPattern
            && x.DateTimeFormat == y.DateTimeFormat
            && x.HorizontalAlignment == y.HorizontalAlignment
            && x.VerticalAlignment == y.VerticalAlignment
            && fontStyleEqualityComparer.Equals(x.FontStyle, y.FontStyle);
    }

    /// <summary>
    /// Returns a hash code for the specified style.
    /// </summary>
    /// <param name="obj">The Style for which a hash code is to be returned.</param>
    /// <returns>A hash code for the specified style.</returns>
    public int GetHashCode(Style obj)
    {
        if (obj is null)
            return 0;

        int hash = (short?)obj.BorderTopColor ?? 0;
        hash += (short?)obj.BorderBottomColor ?? 0;
        hash += (short?)obj.BorderLeftColor ?? 0;
        hash += (short?)obj.BorderRightColor ?? 0;
        hash += (short?)obj.FillForegroundColor ?? 0;
        hash += (short?)obj.FillPattern ?? 0;
        hash += (short?)obj.HorizontalAlignment ?? 0;
        hash += (short?)obj.VerticalAlignment ?? 0;
        hash += obj.DateTimeFormat?.GetHashCode() ?? 0;
        hash += fontStyleEqualityComparer.GetHashCode(obj.FontStyle);

        return hash.GetHashCode();
    }
}
