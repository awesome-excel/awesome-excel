using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.Customization;

public static class ColumnCustomizationExtension
{
    public static ColumnCustomization SetName(this ColumnCustomization columnCustomization, string name)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        columnCustomization.Name = name;
        return columnCustomization;
    }

    public static ColumnCustomization SetStyle(this ColumnCustomization columnCustomization, Action<Style> fn)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);
        InitializeFontStyle(columnCustomization);

        fn(columnCustomization.Style);
        return columnCustomization;
    }

    public static ColumnCustomization SetHorizontalAlignment(this ColumnCustomization columnCustomization, HorizontalAlignment horizontalAlignment)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetHorizontalAlignment(horizontalAlignment);
        return columnCustomization;
    }

    public static ColumnCustomization SetVerticalAlignment(this ColumnCustomization columnCustomization, VerticalAlignment verticalAlignment)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetVerticalAlignment(verticalAlignment);
        return columnCustomization;
    }

    public static ColumnCustomization SetBorderTopColor(this ColumnCustomization columnCustomization, Color color)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetBorderTopColor(color);
        return columnCustomization;
    }

    public static ColumnCustomization SetBorderBottomColor(this ColumnCustomization columnCustomization, Color color)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetBorderBottomColor(color);
        return columnCustomization;
    }

    public static ColumnCustomization SetBorderLeftColor(this ColumnCustomization columnCustomization, Color color)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetBorderLeftColor(color);
        return columnCustomization;
    }

    public static ColumnCustomization SetBorderRightColor(this ColumnCustomization columnCustomization, Color color)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetBorderRightColor(color);
        return columnCustomization;
    }

    public static ColumnCustomization SetFillForegroundColor(this ColumnCustomization columnCustomization, Color color)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetFillForegroundColor(color);
        return columnCustomization;
    }

    public static ColumnCustomization SetFontName(this ColumnCustomization colucolumnCustomizationnInfo, string name)
    {
        if (colucolumnCustomizationnInfo is null)
        {
            throw new ArgumentNullException(nameof(colucolumnCustomizationnInfo));
        }

        InitializeStyle(colucolumnCustomizationnInfo);
        InitializeFontStyle(colucolumnCustomizationnInfo);

        colucolumnCustomizationnInfo.Style.SetFontName(name);
        return colucolumnCustomizationnInfo;
    }

    public static ColumnCustomization SetFontColor(this ColumnCustomization columnCustomization, Color color)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);
        InitializeFontStyle(columnCustomization);

        columnCustomization.Style.SetFontColor(color);
        return columnCustomization;
    }

    public static ColumnCustomization SetFontHeightInPoints(this ColumnCustomization columnCustomization, short height)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);
        InitializeFontStyle(columnCustomization);

        columnCustomization.Style.SetFontHeightInPoints(height);
        return columnCustomization;
    }

    public static ColumnCustomization SetFontBold(this ColumnCustomization columnCustomization, bool isBold)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);
        InitializeFontStyle(columnCustomization);

        columnCustomization.Style.SetFontBold(isBold);
        return columnCustomization;
    }

    public static ColumnCustomization SetDateTimeFormat(this ColumnCustomization columnCustomization, string format)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        InitializeStyle(columnCustomization);

        columnCustomization.Style.SetDateTimeFormat(format);
        return columnCustomization;
    }

    private static void InitializeStyle(ColumnCustomization columnCustomization)
    {
        if (columnCustomization.Style is null)
        {
            columnCustomization.Style = new();
        }
    }

    private static void InitializeFontStyle(ColumnCustomization columnCustomization)
    {
        if (columnCustomization.Style.FontStyle is null)
        {
            columnCustomization.Style.FontStyle = new();
        }
    }

    public static ColumnCustomization Exclude(this ColumnCustomization columnCustomization)
    {
        if (columnCustomization is null)
        {
            throw new ArgumentNullException(nameof(columnCustomization));
        }

        columnCustomization.Excluded = true;
        return columnCustomization;
    }
}
