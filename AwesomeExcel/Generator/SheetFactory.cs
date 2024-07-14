using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;
using AwesomeExcel.Customization.Services;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("Tests")]

namespace AwesomeExcel.Generator;

/// <summary>
/// Defines a factory of objects of type Sheet.
/// </summary>
internal class SheetFactory
{
    /// <summary>
    /// Create a new Sheet object.
    /// </summary>
    /// <typeparam name="TSheet">The type of the rows of the sheet.</typeparam>
    /// <param name="rows">The rows of the sheet.</param>
    /// <param name="sheetCustomization">Additional informations used to customize the sheet.</param>
    /// <param name="columnCustomizationService">Additional informations used to customize the columns of the sheet.</param>
    /// <returns>An Excel Sheet with the given rows and customizations.</returns>
    /// <exception cref="ArgumentNullException">rows is null</exception>
    /// <exception cref="InvalidOperationException">rows contains null elements</exception>
    public Sheet Create<TSheet>(
        IEnumerable<TSheet> rows,
        SheetCustomization<TSheet> sheetCustomization,
        IReadOnlyDictionary<PropertyInfo, ColumnCustomization> columnsCustomization,
        IReadOnlyDictionary<PropertyInfo, CellCustomization> cellsCustomization)
    {
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }

        if (rows.Any(r => r is null))
        {
            throw new InvalidOperationException(nameof(rows));
        }

        PropertyInfo[] properties = typeof(TSheet).GetProperties();
        List<Row> sheetRows = GetRows(rows, properties, cellsCustomization);
        List<Column> sheetColumns = GetColumns(properties, columnsCustomization);

        return new Sheet
        {
            Name = sheetCustomization?.Name,
            Rows = sheetRows,
            Columns = sheetColumns,
            HasHeader = sheetCustomization?.HasHeader ?? false,
            Style = sheetCustomization?.Style,
            HeaderStyle = sheetCustomization?.HeaderStyle,
            IsReadOnly = sheetCustomization?.IsReadOnly ?? false
        };
    }

    private List<Row> GetRows(IEnumerable rows, IEnumerable<PropertyInfo> properties, IReadOnlyDictionary<PropertyInfo, CellCustomization> cellsCustomization)
    {
        List<Row> sheetRows = new();

        foreach (object row in rows)
        {
            Row excelRow = GetRow(row, properties, cellsCustomization);
            sheetRows.Add(excelRow);
        }

        return sheetRows;
    }

    private Row GetRow(object row, IEnumerable<PropertyInfo> properties, IReadOnlyDictionary<PropertyInfo, CellCustomization> cellsCustomization)
    {
        IList<Cell> cells = properties
            .Select(pi =>
            {
                object pValue = pi.GetValue(row, null);
                CellCustomization cellCustomization = null;
                cellsCustomization?.TryGetValue(pi, out cellCustomization);

                if (cellCustomization is null)
                {
                    return new Cell
                    {
                        Value = pValue,
                        Style = null
                    };
                }
                else
                {
                    var resolver = new CellStyleCustomizationResolver(cellCustomization, pValue);
                    Style s = resolver.Resolve();
                    return new Cell
                    {
                        Value = pValue,
                        Style = s
                    };
                }
            })
            .ToList();

        return new Row
        {
            Cells = cells
        };
    }

    private List<Column> GetColumns(IEnumerable<PropertyInfo> properties, IReadOnlyDictionary<PropertyInfo, ColumnCustomization> columnsCustomization)
    {
        IEnumerable<PropertyInfo> filteredProperties;

        if (columnsCustomization is null)
        {
            filteredProperties = properties;
        }
        else
        {
            filteredProperties = properties.Where(pi =>
            {
                bool succeed = columnsCustomization.TryGetValue(pi, out ColumnCustomization value);

                if (!succeed || value is null)
                    return true;

                return value.Excluded == false;
            });
        }

        List<Column> excelColumns = filteredProperties
            .Select(pi =>
            {
                ColumnCustomization customizxation = null;
                columnsCustomization?.TryGetValue(pi, out customizxation);

                string columnName = customizxation?.Name ?? pi.Name;
                Style columnStyle = customizxation?.Style;

                var excelColumnType = GetColumnType(pi.PropertyType);

                return new Column()
                {
                    Name = columnName,
                    ColumnType = excelColumnType,
                    Style = columnStyle
                };
            })
            .ToList();

        return excelColumns;
    }

    private ColumnType GetColumnType(Type type)
    {
        Dictionary<Type, ColumnType> conversionTable = new()
        {
            { typeof(char), ColumnType.String },
            { typeof(string), ColumnType.String },
            { typeof(DateTime), ColumnType.DateTime },
            { typeof(DateTimeOffset), ColumnType.DateTime },
            { typeof(DateOnly), ColumnType.DateTime },
            { typeof(TimeOnly), ColumnType.DateTime },
            { typeof(bool), ColumnType.Numeric },
            { typeof(byte), ColumnType.Numeric },
            { typeof(short), ColumnType.Numeric },
            { typeof(ushort), ColumnType.Numeric },
            { typeof(int), ColumnType.Numeric },
            { typeof(uint), ColumnType.Numeric },
            { typeof(long), ColumnType.Numeric },
            { typeof(ulong), ColumnType.Numeric },
            { typeof(float), ColumnType.Numeric },
            { typeof(double), ColumnType.Numeric },
            { typeof(decimal), ColumnType.Numeric }
        };

        type = Nullable.GetUnderlyingType(type) ?? type;

        if (conversionTable.TryGetValue(type, out ColumnType value))
        {
            return value;
        }
        else if (type.IsClass)
        {
            throw new InvalidOperationException();
        }
        else if (type.IsEnum)
        {
            return ColumnType.String;
        }
        else if (type.IsValueType)
        {
            return ColumnType.String;
        }
        else
        {
            throw new NotSupportedException();
        }
    }
}
