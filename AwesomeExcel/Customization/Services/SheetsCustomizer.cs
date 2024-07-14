using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using System.Linq.Expressions;
using System.Reflection;

namespace AwesomeExcel.Customization.Services;

public abstract class SheetsCustomizer
{
    public WorkbookCustomization Workbook { get; } = new();

    private Dictionary<SheetCustomization, ColumnsCustomizer> dict = new();
    private Dictionary<SheetCustomization, CellsCustomizer> cells = new();

    public ColumnCustomization Column<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        if (!dict.ContainsKey(sheet))
        {
            dict.Add(sheet, new ColumnsCustomizer());
        }

        var customizer = dict[sheet];

        MemberExpression me = selector.Body as MemberExpression;
        PropertyInfo pi = me.Member as PropertyInfo;

        return customizer.GetOrCreateColumn(pi);
    }

    public CellCustomization<TProperty> Cells<TSheet, TProperty>(SheetCustomization<TSheet> sheet, Expression<Func<TSheet, TProperty>> selector)
    {
        return null;
    }

    public IReadOnlyDictionary<PropertyInfo, ColumnCustomization> GetCustomizedColumns<T>(SheetCustomization<T> sheet)
    {
        if (!dict.ContainsKey(sheet))
            return null;

        IReadOnlyDictionary<PropertyInfo, ColumnCustomization> ccs = dict[sheet].GetCustomizedColumn();
        return ccs;
    }

    public IReadOnlyDictionary<PropertyInfo, CellCustomization> GetCustomizedCells<T>(SheetCustomization<T> sheet)
    {
        if (!cells.ContainsKey(sheet))
            return null;

        IReadOnlyDictionary<PropertyInfo, CellCustomization> ccs = cells[sheet].GetCustomizedCells();
        return ccs;
    }
}

public class SheetsCustomizer<TSheet1, TSheet2> : SheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
}

public class SheetsCustomizer<TSheet1, TSheet2, TSheet3> : SheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
}

public class SheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4> : SheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
    public SheetCustomization<TSheet4> Sheet4 { get; } = new();
}

public class SheetsCustomizer<TSheet1, TSheet2, TSheet3, TSheet4, TSheet5> : SheetsCustomizer
{
    public SheetCustomization<TSheet1> Sheet1 { get; } = new();
    public SheetCustomization<TSheet2> Sheet2 { get; } = new();
    public SheetCustomization<TSheet3> Sheet3 { get; } = new();
    public SheetCustomization<TSheet4> Sheet4 { get; } = new();
    public SheetCustomization<TSheet5> Sheet5 { get; } = new();
}
