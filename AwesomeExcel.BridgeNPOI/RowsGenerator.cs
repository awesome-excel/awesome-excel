using AwesomeExcel.Core.Services;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

internal class RowsGenerator
{
    private readonly StylesMerger stylesMerger = new();
    private readonly NpoiHelper npoiHelper = new();

    private readonly _NPOI.ISheet npoiSheet;
    private readonly Sheet excelSheet;
    private readonly StyleConverter styleConverter;

    private Style colorBandingEvenRows;
    private Style colorbandingOddRows;

    public RowsGenerator(_NPOI.ISheet npoiSheet, Sheet excelSheet, StyleConverter styleConverter)
    {
        this.npoiSheet = npoiSheet;
        this.excelSheet = excelSheet;
        this.styleConverter = styleConverter;
    }

    public void GenerateHeaderRow()
    {
        if (!excelSheet.HasHeader)
            return;

        if (excelSheet.Columns is null)
            return;

        Style? style = stylesMerger.Merge(excelSheet.Style, excelSheet.HeaderStyle);
        _NPOI.ICellStyle headerStyle = styleConverter.Convert(style);
        IEnumerable<string> columns = excelSheet.Columns.Select(c => c.Name);

        GenerateHeaderRow(npoiSheet, columns, headerStyle);
    }

    private void GenerateHeaderRow(_NPOI.ISheet sheet, IEnumerable<string> columnsName, _NPOI.ICellStyle headerStyle)
    {
        _NPOI.IRow headerRow = sheet.CreateRow(0);

        // The header row needs to be taller than normal rows
        headerRow.HeightInPoints *= 1.3f;

        int columnIndex = 0;
        foreach (string columnName in columnsName)
        {
            _NPOI.ICell cell = CreateCell(headerRow, columnIndex, _NPOI.CellType.String, headerStyle);
            cell.SetCellValue(columnName);
            columnIndex++;
        }
    }

    public void GenerateRows()
    {
        int rowsCount = excelSheet.Rows?.Count ?? 0;

        for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
        {
            int rowNumber = rowIndex + (excelSheet.HasHeader ? 1 : 0);

            Row? row = excelSheet.Rows[rowIndex];

            bool skipNullRows = false;
            if (row == null && skipNullRows)
            {
                continue;
            }

            Style? rowStyle = stylesMerger.Merge(row?.Style);

            _NPOI.ICellStyle? npoiStyle = styleConverter.Convert(rowStyle);
            _NPOI.IRow npoiRow = CreateRow(npoiSheet, rowNumber, npoiStyle);

            GenerateCells(npoiRow, row, rowNumber);
        }
    }

    public _NPOI.IRow CreateRow(_NPOI.ISheet sheet, int rowNumber, _NPOI.ICellStyle? rowStyle)
    {
        _NPOI.IRow row = sheet.CreateRow(rowNumber);

        if (rowStyle != null)
        {
            row.RowStyle = rowStyle;
        }

        return row;
    }

    private void GenerateCells(_NPOI.IRow npoiRow, Row? excelRow, int rowNumber)
    {
        int rowColumnsCount = excelRow?.Cells?.Count ?? 0;
        int sheetColumnCount = excelSheet.Columns?.Count ?? 0;

        for (int columnIndex = 0; columnIndex < rowColumnsCount; columnIndex++)
        {
            Column column;

            if (columnIndex < sheetColumnCount)
            {
                column = excelSheet.Columns[columnIndex];
            }
            else
            {
                // Fill the missing cells of the given row with a blank cell
                column = new()
                {
                    ColumnType = ColumnType.String,
                    Name = null,
                    Style = null
                };
            }

            Cell cell = excelRow.Cells[columnIndex];
            Style? colorBanding = GetColorBanding(rowNumber);
            Style? dateTimeFormat = GetDefaultDateTimeFormat(column.ColumnType);
            Style? style = stylesMerger.Merge(dateTimeFormat, excelSheet.Style, colorBanding, column.Style, cell?.Style);

            _NPOI.CellType cellType = GetCellType(column.ColumnType);
            _NPOI.ICellStyle? npoiStyle = styleConverter.Convert(style);
            _NPOI.ICell npoiCell = CreateCell(npoiRow, columnIndex, cellType, npoiStyle);

            npoiHelper.SetCellValue(npoiCell, column.ColumnType, cell?.Value);
        }
    }

    private _NPOI.CellType GetCellType(ColumnType columnType) => columnType switch
    {
        ColumnType.Numeric => _NPOI.CellType.Numeric,
        ColumnType.String => _NPOI.CellType.String,
        ColumnType.DateTime => _NPOI.CellType.Numeric,
        _ => _NPOI.CellType.String,
    };

    private _NPOI.ICell CreateCell(_NPOI.IRow row, int columnIndex, _NPOI.CellType cellType, _NPOI.ICellStyle? cellStyle)
    {
        _NPOI.ICell cell = row.CreateCell(columnIndex, cellType);

        if (cellStyle != null)
        {
            cell.CellStyle = cellStyle;
        }

        return cell;
    }

    private static Style? GetDefaultDateTimeFormat(ColumnType columnType)
    {
        if (columnType != ColumnType.DateTime)
            return null;

        return new Style
        {
            DateTimeFormat = "yyyy/mm/dd"
        };
    }

    private Style? GetColorBanding(int rowNumber)
    {
        ColorBanding colorBanding = excelSheet.Style?.ColorBanding;

        if (colorBanding is null)
            return null;

        bool isEven = rowNumber % 2 == 0;

        if (isEven && colorBandingEvenRows != null)
        {
            // Reuse the style I already created
            return colorBandingEvenRows;
        }

        if (!isEven && colorbandingOddRows != null)
        {
            // Reuse the style I already created
            return colorbandingOddRows;
        }

        Color color = isEven ? colorBanding.EvenRows : colorBanding.OddRows;

        var s = new Style()
        {
            FillForegroundColor = color,
            FillPattern = FillPattern.SolidForeground
        };

        if (isEven)
        {
            // Save the style for reuse
            colorBandingEvenRows = s;
        }
        else
        {
            // Save the style for reuse
            colorbandingOddRows = s;
        }

        return s;
    }
}
