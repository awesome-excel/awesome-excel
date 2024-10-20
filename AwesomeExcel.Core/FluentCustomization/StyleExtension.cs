using AwesomeExcel.Models;

namespace AwesomeExcel;

public static class StyleExtension
{
    public static Style SetHorizontalAlignment(this Style style, HorizontalAlignment horizontalAlignment)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.HorizontalAlignment = horizontalAlignment;
        return style;
    }

    public static Style SetVerticalAlignment(this Style style, VerticalAlignment verticalAlignment)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.VerticalAlignment = verticalAlignment;
        return style;
    }

    public static Style SetBorderTopColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.BorderTopColor = color;
        return style;
    }

    public static Style SetBorderBottomColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.BorderBottomColor = color;
        return style;
    }

    public static Style SetBorderLeftColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.BorderLeftColor = color;
        return style;
    }

    public static Style SetBorderRightColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.BorderRightColor = color;
        return style;
    }

    public static Style SetBordersColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        if (!style.BorderTopColor.HasValue) style.BorderTopColor = color;
        if (!style.BorderBottomColor.HasValue) style.BorderBottomColor = color;
        if (!style.BorderLeftColor.HasValue) style.BorderLeftColor = color;
        if (!style.BorderRightColor.HasValue) style.BorderRightColor = color;

        return style;
    }

    public static Style SetFillForegroundColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.FillForegroundColor = color;
        return style;
    }

    public static Style SetFontName(this Style style, string name)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.FontStyle.Name = name;
        return style;
    }

    public static Style SetFontColor(this Style style, Color color)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.FontStyle.Color = color;
        return style;
    }

    public static Style SetFontHeightInPoints(this Style style, short height)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.FontStyle.HeightInPoints = height;
        return style;
    }

    public static Style SetFontBold(this Style style, bool isBold)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.FontStyle.IsBold = isBold;
        return style;
    }

    public static Style SetDateTimeFormat(this Style style, string format)
    {
        if (style is null)
        {
            throw new ArgumentNullException(nameof(style));
        }

        style.DateTimeFormat = format;
        return style;
    }
}