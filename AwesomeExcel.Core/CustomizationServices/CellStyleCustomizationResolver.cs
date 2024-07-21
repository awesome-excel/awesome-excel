using System.Reflection;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Tests")]

namespace AwesomeExcel.Core.CustomizationServices;

internal class CellStyleCustomizationResolver
{
    private readonly CellStyleCustomization customization;
    private readonly object cellValue;

    public CellStyleCustomizationResolver(CellStyleCustomization customization, object cellValue)
    {
        this.customization = customization;
        this.cellValue = cellValue;
    }

    public Style? Resolve()
    {
        Color? borderTopColor = GetBorderTopColor(customization);
        Color? borderBottomColor = GetBorderBottomColor(customization);
        Color? borderLeftColor = GetBorderLeftColor(customization);
        Color? borderRightColor = GetBorderRightColor(customization);
        Color? fillForegroundColor = GetFillForegroundColor(customization);
        FillPattern? fillPattern = GetFillPattern(customization);
        string dateTimeFormat = GetDateTimeFormat(customization);
        HorizontalAlignment? horizontalAlignment = GetHorizontalAlignment(customization);
        VerticalAlignment? verticalAlignment = GetVerticalAlignment(customization);
        FontStyle fontStyle = GetFontStyle(customization);

        return new Style
        {
            BorderTopColor = borderTopColor,
            BorderBottomColor = borderBottomColor,
            BorderLeftColor = borderLeftColor,
            BorderRightColor = borderRightColor,
            FillForegroundColor = fillForegroundColor,
            FillPattern = fillPattern,
            DateTimeFormat = dateTimeFormat,
            HorizontalAlignment = horizontalAlignment,
            VerticalAlignment = verticalAlignment,
            FontStyle = fontStyle
        };
    }

    private FontStyle GetFontStyle(CellStyleCustomization sc)
    {
        CellFontStyleCustomization cfsc = GetFontStyleCustomization(sc);

        if (cfsc is null)
            return null;

        Color? color = GetFontColor(cfsc);
        short? heightInPoints = GetFontHeightInPoints(cfsc);
        bool? isBold = GetFontBold(cfsc);
        string name = GetFontName(cfsc);

        return new FontStyle
        {
            Color = color,
            HeightInPoints = heightInPoints,
            IsBold = isBold,
            Name = name
        };
    }

    private Color? GetBorderTopColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderTopColor);
        return GetValue<Color?>(sc, pName, cellValue);
    }

    private Color? GetBorderBottomColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderBottomColor);
        return GetValue<Color?>(sc, pName, cellValue);
    }

    private Color? GetBorderLeftColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderLeftColor);
        return GetValue<Color?>(sc, pName, cellValue);
    }

    private Color? GetBorderRightColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderRightColor);
        return GetValue<Color?>(sc, pName, cellValue);
    }

    private Color? GetFillForegroundColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.FillForegroundColor);
        return GetValue<Color?>(sc, pName, cellValue);
    }

    private FillPattern? GetFillPattern(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.FillPattern);
        return GetValue<FillPattern?>(sc, pName, cellValue);
    }

    private HorizontalAlignment? GetHorizontalAlignment(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.HorizontalAlignment);
        return GetValue<HorizontalAlignment?>(sc, pName, cellValue);
    }

    private VerticalAlignment? GetVerticalAlignment(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.VerticalAlignment);
        return GetValue<VerticalAlignment?>(sc, pName, cellValue);
    }

    private string GetDateTimeFormat(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.DateTimeFormat);
        return GetValue<string>(sc, pName, cellValue);
    }

    private CellFontStyleCustomization GetFontStyleCustomization(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.FontStyle);

        Type type = sc.GetType();
        PropertyInfo pi = type.GetProperty(pName);
        var pValue = (CellFontStyleCustomization)pi.GetValue(sc);

        return pValue;
    }

    private Color? GetFontColor(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.Color);
        return GetFontValue<Color?>(cfsc, pName, cellValue);
    }

    private short? GetFontHeightInPoints(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.HeightInPoints);
        return GetFontValue<short?>(cfsc, pName, cellValue);
    }

    private bool? GetFontBold(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.IsBold);
        return GetFontValue<bool?>(cfsc, pName, cellValue);
    }

    private string GetFontName(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.Name);
        return GetFontValue<string>(cfsc, pName, cellValue);
    }

    private T GetFontValue<T>(CellFontStyleCustomization cfsc, string pName, object value)
    {
        Type fscType = cfsc.GetType();
        PropertyInfo pi = fscType.GetProperty(pName);
        var pValue = (Delegate)pi.GetValue(cfsc);
        var result = Invoke<T>(pValue, value);
        return result;
    }

    private T GetValue<T>(CellStyleCustomization sc, string pName, object cellValue)
    {
        Type type = sc.GetType();
        PropertyInfo pi = type.GetProperty(pName);
        var pValue = (Delegate)pi.GetValue(sc);
        var result = Invoke<T>(pValue, cellValue);
        return result;
    }

    private static T Invoke<T>(Delegate fn, object value)
    {
        return (T)fn?.Method.Invoke(fn.Target, new[] { value });
    }
}
