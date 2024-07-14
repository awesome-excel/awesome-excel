using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class RowsGenerator
{
    private readonly NpoiFacade npoiFacade = new();
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

        _Excel.Style excelStyle = stylesMerger.Merge(excelSheet.Style, excelSheet.HeaderStyle);
        _NPOI.ICellStyle npoiStyle = styleConverter.Convert(excelStyle);
        IEnumerable<string> columns = excelSheet.Columns.Select(c => c.Name);

        npoiFacade.CreateHeaderRow(npoiSheet, npoiStyle, columns);
    }

    public void GenerateRows()
    {
        int rowsCount = excelSheet.Rows?.Count ?? 0;

        for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
        {
            int rowNumber = rowIndex + (excelSheet.HasHeader ? 1 : 0);

            _Excel.Row row = excelSheet.Rows[rowIndex];
            _Excel.Style rowStyle = stylesMerger.Merge(row?.Style);

            _NPOI.ICellStyle npoiStyle = styleConverter.Convert(rowStyle);
            _NPOI.IRow npoiRow = npoiFacade.CreateRow(npoiSheet, rowNumber, npoiStyle);

            GenerateCells(npoiRow, row, rowNumber);
        }
    }

    private _Excel.Style GetColorBandingStyle(int rowNumber)
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

    private void GenerateCells(_NPOI.IRow npoiRow, _Excel.Row excelRow, int rowNumber)
    {
        int rowColumnsCount = excelRow?.Cells?.Count ?? 0;
        int sheetColumnCount = (excelSheet.Columns?.Count ?? 0);

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
                column = GetNewEmptyCell();
            }

            _Excel.Cell cell = excelRow.Cells[columnIndex];
            _Excel.Style colorBandingStyle = GetColorBandingStyle(rowNumber);
            _Excel.Style dateTimeFormatStyle = GetDateTimeFormatStyle(column.ColumnType);
            _Excel.Style style = stylesMerger.Merge(dateTimeFormatStyle, excelSheet.Style, colorBandingStyle, column.Style, cell?.Style);

            _NPOI.CellType cellType = npoiHelper.GetCellType(column.ColumnType);
            _NPOI.ICellStyle npoiStyle = styleConverter.Convert(style);
            _NPOI.ICell npoiCell = npoiFacade.CreateCell(npoiRow, columnIndex, cellType, npoiStyle);

            npoiHelper.SetCellValue(npoiCell, column.ColumnType, cell?.Value);
        }
    }

    private static _Excel.Column GetNewEmptyCell()
    {
        return new()
        {
            ColumnType = _Excel.ColumnType.String,
            Name = null,
            Style = null
        };
    }

    private static _Excel.Style GetDateTimeFormatStyle(_Excel.ColumnType columnType)
    {
        if (columnType != _Excel.ColumnType.DateTime)
            return null;

        return new _Excel.Style
        {
            DateTimeFormat = "yyyy/mm/dd"
        };
    }
}
