using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public class SheetCustomizer<T>
{
    private readonly ColumnsCustomizer ccs = new();
    private readonly CellsCustomizer<T> cells = new();

    public WorkbookCustomization Workbook { get; } = new();

    public SheetCustomization<T> Sheet { get; } = new();

    public ColumnCustomization Column<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        var x = ccs.GetOrCreateColumn(pi);
        return x;
    }

    public CellCustomization<TProperty> Cells<TProperty>(Expression<Func<T, TProperty>> selector)
    {
        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        return (CellCustomization<TProperty>)cells.GetOrCreateCells<TProperty>(pi);
    }

    internal IReadOnlyDictionary<PropertyInfo, ColumnCustomization> GetCustomizedColumns()
    {
        IReadOnlyDictionary<PropertyInfo, ColumnCustomization> x = ccs.GetCustomizedColumn();
        return x;
    }

    internal IReadOnlyDictionary<PropertyInfo, CellCustomization> GetCustomizedCells()
    {
        IReadOnlyDictionary<PropertyInfo, CellCustomization> x = cells.GetCustomizedCells();
        return x;
    }
}
