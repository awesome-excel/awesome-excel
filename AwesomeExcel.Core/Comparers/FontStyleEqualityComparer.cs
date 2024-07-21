namespace AwesomeExcel.Core.Comparers;

/// <summary>
/// Defines methods to support the comparison of objects of type FontStyle for equality.
/// </summary>
public class FontStyleEqualityComparer : IEqualityComparer<FontStyle>
{
    /// <summary>
    /// Determines whether the specified FontStyles are equal.
    /// </summary>
    /// <param name="x">The first font style to compare.</param>
    /// <param name="y">The second font style to compare.</param>
    /// <returns>true if the specified font styles are equal; otherwise, false.</returns>
    public bool Equals(FontStyle x, FontStyle y)
    {
        if (x == y)
            return true;

        if (x is null || y is null)
            return false;

        return x.Name == y.Name
            && x.Color == y.Color
            && x.HeightInPoints == y.HeightInPoints
            && x.IsBold == y.IsBold;
    }

    /// <summary>
    /// Returns a hash code for the specified font style.
    /// </summary>
    /// <param name="obj">The font style for which a hash code is to be returned.</param>
    /// <returns>A hash code for the specified font style.</returns>
    public int GetHashCode(FontStyle obj)
    {
        if (obj is null)
            return 0;

        int hash = 1;
        hash += obj.Name?.GetHashCode() ?? 0;
        hash += (short?)obj.Color ?? 0;
        hash += obj.HeightInPoints ?? 0;
        hash += obj.IsBold.HasValue ? 1 : 0;
        hash += obj.IsBold.HasValue && obj.IsBold.Value ? 2 : 1;

        return hash.GetHashCode();
    }
}
