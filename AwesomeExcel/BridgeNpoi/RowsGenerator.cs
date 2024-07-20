using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class RowsGenerator
{
    private readonly Common.Services.StylesMerger stylesMerger = new();
    private readonly NpoiHelper npoiHelper = new();

    private readonly _NPOI.ISheet npoiSheet;
    private readonly _Excel.Sheet excelSheet;
    private readonly StyleConverter styleConverter;

    private _Excel.Style colorBandingEvenRows;
    private _Excel.Style colorbandingOddRows;

    public RowsGenerator(_NPOI.ISheet npoiSheet, _Excel.Sheet excelSheet, StyleConverter styleConverter)
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

        _Excel.Style? style = stylesMerger.Merge(excelSheet.Style, excelSheet.HeaderStyle);
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

            _Excel.Row? row = excelSheet.Rows[rowIndex];

            bool skipNullRows = false;
            if (row == null && skipNullRows)
            {
                continue;
            }

            _Excel.Style? rowStyle = stylesMerger.Merge(row?.Style);

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

    private void GenerateCells(_NPOI.IRow npoiRow, _Excel.Row? excelRow, int rowNumber)
    {
        int rowColumnsCount = excelRow?.Cells?.Count ?? 0;
        int sheetColumnCount = excelSheet.Columns?.Count ?? 0;

        for (int columnIndex = 0; columnIndex < rowColumnsCount; columnIndex++)
        {
            _Excel.Column column;

            if (columnIndex < sheetColumnCount)
            {
                column = excelSheet.Columns[columnIndex];
            }
            else
            {
                // Fill the missing cells of the given row with a blank cell
                column = new()
                {
                    ColumnType = _Excel.ColumnType.String,
                    Name = null,
                    Style = null
                };
            }

            _Excel.Cell cell = excelRow.Cells[columnIndex];
            _Excel.Style? colorBanding = GetColorBanding(rowNumber);
            _Excel.Style? dateTimeFormat = GetDefaultDateTimeFormat(column.ColumnType);
            _Excel.Style? style = stylesMerger.Merge(dateTimeFormat, excelSheet.Style, colorBanding, column.Style, cell?.Style);

            _NPOI.CellType cellType = GetCellType(column.ColumnType);
            _NPOI.ICellStyle? npoiStyle = styleConverter.Convert(style);
            _NPOI.ICell npoiCell = CreateCell(npoiRow, columnIndex, cellType, npoiStyle);

            npoiHelper.SetCellValue(npoiCell, column.ColumnType, cell?.Value);
        }
    }

    private _NPOI.CellType GetCellType(_Excel.ColumnType columnType) => columnType switch
    {
        _Excel.ColumnType.Numeric => _NPOI.CellType.Numeric,
        _Excel.ColumnType.String => _NPOI.CellType.String,
        _Excel.ColumnType.DateTime => _NPOI.CellType.Numeric,
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

    private static _Excel.Style? GetDefaultDateTimeFormat(_Excel.ColumnType columnType)
    {
        if (columnType != _Excel.ColumnType.DateTime)
            return null;

        return new _Excel.Style
        {
            DateTimeFormat = "yyyy/mm/dd"
        };
    }

    private _Excel.Style? GetColorBanding(int rowNumber)
    {
        _Excel.ColorBanding colorBanding = excelSheet.Style?.ColorBanding;

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

        _Excel.Color color = isEven ? colorBanding.EvenRows : colorBanding.OddRows;

        var s = new _Excel.Style()
        {
            FillForegroundColor = color,
            FillPattern = _Excel.FillPattern.SolidForeground
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
