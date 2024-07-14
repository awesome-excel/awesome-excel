using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.Customization.Fluent;

public static class SheetCustomizationHeaderExtension
{
    public static SheetCustomization SetHeaderHorizontalAlignment(this SheetCustomization sheetCustomization, HorizontalAlignment horizontalAlignment)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetHorizontalAlignment(horizontalAlignment);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderVerticalAlignment(this SheetCustomization sheetCustomization, VerticalAlignment verticalAlignment)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetVerticalAlignment(verticalAlignment);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderBorderTopColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetBorderTopColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderBorderBottomColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetBorderBottomColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderBorderLeftColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetBorderLeftColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderBorderRightColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetBorderRightColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderFillForegroundColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetFillForegroundColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderFontName(this SheetCustomization sheetCustomization, string name)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);
        InitializeHeaderFontStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetFontName(name);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderFontColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);
        InitializeHeaderFontStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetFontColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderFontHeightInPoints(this SheetCustomization sheetCustomization, short height)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);
        InitializeHeaderFontStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetFontHeightInPoints(height);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderFontBold(this SheetCustomization sheetCustomization, bool isBold = true)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);
        InitializeHeaderFontStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetFontBold(isBold);
        return sheetCustomization;
    }

    public static SheetCustomization SetHeaderDateTimeFormat(this SheetCustomization sheetCustomization, string format)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeHeaderStyle(sheetCustomization);
        InitializeHeaderFontStyle(sheetCustomization);

        sheetCustomization.HeaderStyle.SetDateTimeFormat(format);
        return sheetCustomization;
    }

    private static void InitializeHeaderStyle(SheetCustomization sheetCustomization)
    {
        if (sheetCustomization.HeaderStyle is null)
        {
            sheetCustomization.HeaderStyle = new();
        }
    }

    private static void InitializeHeaderFontStyle(SheetCustomization sheetCustomization)
    {
        if (sheetCustomization.HeaderStyle.FontStyle is null)
        {
            sheetCustomization.HeaderStyle.FontStyle = new();
        }
    }
}