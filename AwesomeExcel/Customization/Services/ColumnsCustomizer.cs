using AwesomeExcel.Customization.Models;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

internal class ColumnsCustomizer 
{
    private readonly Dictionary<PropertyInfo, ColumnCustomization> customizedColumns = new();

    public ColumnCustomization GetColumn(PropertyInfo pi)
    {
        customizedColumns.TryGetValue(pi, out ColumnCustomization value);
        return value;
    }

    public ColumnCustomization GetOrCreateColumn(PropertyInfo pi)
    {
        if (customizedColumns.TryGetValue(pi, out ColumnCustomization value))
        {
            return value;
        }
        else
        {
            ColumnCustomization cc = new ColumnCustomization();
            customizedColumns.Add(pi, cc);
            return cc;
        }
    }

    public IReadOnlyDictionary<PropertyInfo, ColumnCustomization> GetCustomizedColumn()
    {
        return customizedColumns;
    }
}
