using AwesomeExcel.Customization.Models;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

internal class CellsCustomizer
{
    private readonly Dictionary<PropertyInfo, CellCustomization> customizedCells = new();

    public CellCustomization GetCells(PropertyInfo pi)
    {
        customizedCells.TryGetValue(pi, out CellCustomization value);
        return value;
    }

    public CellCustomization GetOrCreateCells<T>(PropertyInfo pi)
    {
        if (customizedCells.TryGetValue(pi, out CellCustomization value))
        {
            return value;
        }
        else
        {
            CellCustomization<T> ci = new();
            customizedCells.Add(pi, ci);
            return ci;
        }
    }

    public IReadOnlyDictionary<PropertyInfo, CellCustomization> GetCustomizedCells()
    {
        return customizedCells;
    }
}

internal class CellsCustomizer<T> : CellsCustomizer { }