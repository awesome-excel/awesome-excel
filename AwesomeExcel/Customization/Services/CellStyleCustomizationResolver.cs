using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public class CellStyleCustomizationResolver
{
    private readonly CellCustomization customization;
    private readonly object value;

    public CellStyleCustomizationResolver(CellCustomization customization, object value)
    {
        this.customization = customization;
        this.value = value;
    }

    public Style Resolve()
    {
        CellStyleCustomization csc = GetCellStyleCustomization(customization);

        if (csc is null)
            return null;

        Color? borderTopColor = GetBorderTopColor(csc);
        Color? borderBottomColor = GetBorderBottomColor(csc);
        Color? borderLeftColor = GetBorderLeftColor(csc);
        Color? borderRightColor = GetBorderRightColor(csc);
        Color? fillForegroundColor = GetFillForegroundColor(csc);
        FillPattern? fillPattern = GetFillPattern(csc);
        string dateTimeFormat = GetDateTimeFormat(csc);
        HorizontalAlignment? horizontalAlignment = GetHorizontalAlignment(csc);
        VerticalAlignment? verticalAlignment = GetVerticalAlignment(csc);
        FontStyle fontStyle = GetFontStyle(csc);

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

    private CellStyleCustomization GetCellStyleCustomization(CellCustomization customization)
    {
        string pName = nameof(CellCustomization<object>.Style);
        Type type = customization.GetType();
        PropertyInfo pi = type.GetProperty(pName);
        var csc = (CellStyleCustomization)pi.GetValue(customization);
        return csc;
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
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetBorderBottomColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderBottomColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetBorderLeftColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderLeftColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetBorderRightColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.BorderRightColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private Color? GetFillForegroundColor(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.FillForegroundColor);
        return GetValue<Color?>(sc, pName, value);
    }

    private FillPattern? GetFillPattern(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.FillPattern);
        return GetValue<FillPattern?>(sc, pName, value);
    }

    private HorizontalAlignment? GetHorizontalAlignment(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.HorizontalAlignment);
        return GetValue<HorizontalAlignment?>(sc, pName, value);
    }

    private VerticalAlignment? GetVerticalAlignment(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.VerticalAlignment);
        return GetValue<VerticalAlignment?>(sc, pName, value);
    }

    private string GetDateTimeFormat(CellStyleCustomization sc)
    {
        const string pName = nameof(CellStyleCustomization<object>.DateTimeFormat);
        return GetValue<string>(sc, pName, value);
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
        return GetFontValue<Color?>(cfsc, pName, value);
    }

    private short? GetFontHeightInPoints(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.HeightInPoints);
        return GetFontValue<short?>(cfsc, pName, value);
    }

    private bool? GetFontBold(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.IsBold);
        return GetFontValue<bool?>(cfsc, pName, value);
    }

    private string GetFontName(CellFontStyleCustomization cfsc)
    {
        const string pName = nameof(CellFontStyleCustomization<object>.Name);
        return GetFontValue<string>(cfsc, pName, value);
    }

    private T GetFontValue<T>(CellFontStyleCustomization cfsc, string pName, object value)
    {
        Type fscType = cfsc.GetType();
        PropertyInfo pi = fscType.GetProperty(pName);
        var pValue = (Delegate)pi.GetValue(cfsc);
        var result = Invoke<T>(pValue, value);
        return result;
    }

    private T GetValue<T>(CellStyleCustomization sc, string pName, object value)
    {
        Type type = sc.GetType();
        PropertyInfo pi = type.GetProperty(pName);
        var pValue = (Delegate)pi.GetValue(sc);
        var result = Invoke<T>(pValue, value);
        return result;
    }

    private static T Invoke<T>(Delegate fn, object value)
    {
        return (T)fn?.Method.Invoke(fn.Target, new[] { value });
    }
}
