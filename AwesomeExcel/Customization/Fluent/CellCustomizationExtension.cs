using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.Customization;

public static class CellCustomizationExtension
{
    public static CellCustomization<T> SetHorizontalAlignment<T>(this CellCustomization<T> cellCustomizatioon, Func<T, HorizontalAlignment?> horizontalAlignment)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.HorizontalAlignment = horizontalAlignment;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetVerticalAlignment<T>(this CellCustomization<T> cellCustomizatioon, Func<T, VerticalAlignment?> verticalAlignment)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.VerticalAlignment = verticalAlignment;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetBorderTopColor<T>(this CellCustomization<T> cellCustomizatioon, Func<T, Color?> color)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.BorderTopColor = color;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetBorderBottomColor<T>(this CellCustomization<T> cellCustomizatioon, Func<T, Color?> color)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.BorderBottomColor = color;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetBorderLeftColor<T>(this CellCustomization<T> cellCustomizatioon, Func<T, Color?> color)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.BorderLeftColor = color;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetBorderRightColor<T>(this CellCustomization<T> cellCustomizatioon, Func<T, Color?> color)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.BorderRightColor = color;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetFillForegroundColor<T>(this CellCustomization<T> cellCustomizatioon, Func<T, Color?> color)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.FillForegroundColor = color;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetFontName<T>(this CellCustomization<T> cellCustomizatioon, Func<T, string> name)
    {
        InitializeStyle(cellCustomizatioon);
        InitializeFontStyle(cellCustomizatioon);

        cellCustomizatioon.Style.FontStyle.Name = name;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetFontColor<T>(this CellCustomization<T> cellCustomizatioon, Func<T, Color?> color)
    {
        InitializeStyle(cellCustomizatioon);
        InitializeFontStyle(cellCustomizatioon);

        cellCustomizatioon.Style.FontStyle.Color = color;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetFontHeightInPoints<T>(this CellCustomization<T> cellCustomizatioon, Func<T, short?> height)
    {
        InitializeStyle(cellCustomizatioon);
        InitializeFontStyle(cellCustomizatioon);

        cellCustomizatioon.Style.FontStyle.HeightInPoints = height;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetFontBold<T>(this CellCustomization<T> cellCustomizatioon, Func<T, bool?> isBold)
    {
        InitializeStyle(cellCustomizatioon);
        InitializeFontStyle(cellCustomizatioon);

        cellCustomizatioon.Style.FontStyle.IsBold = isBold;
        return cellCustomizatioon;
    }

    public static CellCustomization<T> SetDateTimeFormat<T>(this CellCustomization<T> cellCustomizatioon, Func<T, string> format)
    {
        InitializeStyle(cellCustomizatioon);

        cellCustomizatioon.Style.DateTimeFormat = format;
        return cellCustomizatioon;
    }

    private static void InitializeStyle<T>(CellCustomization<T> cellCustomizatioon)
    {
        if (cellCustomizatioon.Style is null)
        {
            cellCustomizatioon.Style = new();
        }
    }

    private static void InitializeFontStyle<T>(CellCustomization<T> cellCustomizatioon)
    {
        if (cellCustomizatioon.Style.FontStyle is null)
        {
            cellCustomizatioon.Style.FontStyle = new();
        }
    }
}
