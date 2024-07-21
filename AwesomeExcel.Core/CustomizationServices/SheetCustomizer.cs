using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel;

public class SheetCustomizer<T> : SheetCustomization<T>
{
    private readonly Dictionary<PropertyInfo, ICellCustomization> customizedCells = new();
    private readonly Dictionary<PropertyInfo, ColumnCustomization> customizedColumns = new();

    public WorkbookCustomization Workbook { get; } = new();

    public ColumnCustomization Column<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        ColumnCustomization cc = GetOrCreateColumn(pi);
        return cc;
    }

    public CellCustomization<TProperty> Cells<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        CellCustomization<TProperty> cc = GetOrCreateCells<TProperty>(pi);
        return cc;
    }

    internal IReadOnlyDictionary<PropertyInfo, ColumnCustomization> GetColumns()
    {
        return customizedColumns;
    }

    internal IReadOnlyDictionary<PropertyInfo, ICellCustomization> GetCells()
    {
        return customizedCells;
    }

    private CellCustomization<TProperty> GetOrCreateCells<TProperty>(PropertyInfo pi)
    {
        if (customizedCells.TryGetValue(pi, out ICellCustomization value))
        {
            return (CellCustomization<TProperty>)value;
        }

        CellCustomization<TProperty> ci = new();
        customizedCells.Add(pi, ci);
        return ci;
    }

    private ColumnCustomization GetOrCreateColumn(PropertyInfo pi)
    {
        if (customizedColumns.TryGetValue(pi, out ColumnCustomization value))
        {
            return value;
        }

        ColumnCustomization cc = new();
        customizedColumns.Add(pi, cc);
        return cc;
    }
}
