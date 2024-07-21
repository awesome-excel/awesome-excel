namespace AwesomeExcel;

public static class SheetCustomizationStyleExtension
{
    public static SheetCustomization SetHorizontalAlignment(this SheetCustomization sheetCustomization, HorizontalAlignment horizontalAlignment)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetHorizontalAlignment(horizontalAlignment);
        return sheetCustomization;
    }

    public static SheetCustomization SetVerticalAlignment(this SheetCustomization sheetCustomization, VerticalAlignment verticalAlignment)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetVerticalAlignment(verticalAlignment);
        return sheetCustomization;
    }

    public static SheetCustomization SetBorderTopColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetBorderTopColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetBorderBottomColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetBorderBottomColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetBorderLeftColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetBorderLeftColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetBorderRightColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetBorderRightColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetBordersColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetBordersColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetFillForegroundColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetFillForegroundColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetFontName(this SheetCustomization sheetCustomization, string name)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);
        InitializeFontStyle(sheetCustomization);

        sheetCustomization.Style.SetFontName(name);
        return sheetCustomization;
    }

    public static SheetCustomization SetFontColor(this SheetCustomization sheetCustomization, Color color)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);
        InitializeFontStyle(sheetCustomization);

        sheetCustomization.Style.SetFontColor(color);
        return sheetCustomization;
    }

    public static SheetCustomization SetFontHeightInPoints(this SheetCustomization sheetCustomization, short height)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);
        InitializeFontStyle(sheetCustomization);

        sheetCustomization.Style.SetFontHeightInPoints(height);
        return sheetCustomization;
    }

    public static SheetCustomization SetFontBold(this SheetCustomization sheetCustomization, bool isBold)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);
        InitializeFontStyle(sheetCustomization);

        sheetCustomization.Style.SetFontBold(isBold);
        return sheetCustomization;
    }

    public static SheetCustomization SetDateTimeFormat(this SheetCustomization sheetCustomization, string format)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        InitializeStyle(sheetCustomization);

        sheetCustomization.Style.SetDateTimeFormat(format);
        return sheetCustomization;
    }


    private static void InitializeStyle(SheetCustomization sheetCustomization)
    {
        if (sheetCustomization.Style is null)
        {
            sheetCustomization.Style = new();
        }
    }

    private static void InitializeFontStyle(SheetCustomization sheetCustomization)
    {
        if (sheetCustomization.Style.FontStyle is null)
        {
            sheetCustomization.Style.FontStyle = new();
        }
    }
}
