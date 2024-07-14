using NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class NpoiFacade
{
    public ICell CreateCell(IRow row, int columnIndex, CellType cellType, ICellStyle cellStyle)
    {
        ICell cell = row.CreateCell(columnIndex, cellType);
        cell.CellStyle = cellStyle;
        return cell;
    }

    public IRow CreateRow(ISheet sheet, int rowNumber, ICellStyle rowStyle)
    {
        IRow row = sheet.CreateRow(rowNumber);
        row.RowStyle = rowStyle;
        return row;
    }

    public void CreateHeaderRow(ISheet sheet, ICellStyle headerStyle, IEnumerable<string> columnsName)
    {
        IRow headerRow = sheet.CreateRow(0);

        // The header row needs to be taller than normal rows
        headerRow.HeightInPoints *= 1.3f;

        int columnIndex = 0;
        foreach (string columnName in columnsName)
        {
            ICell cell = CreateCell(headerRow, columnIndex, CellType.String, headerStyle);
            cell.SetCellValue(columnName);
            columnIndex++;
        }
    }

    public void AutoSizeColumns(ISheet sheet, int columnsCount)
    {
        for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
        {
            sheet.AutoSizeColumn(columnIndex);

            // The column width is still small after the autosize, I prefer it to be a little bit more wider
            int widthAfterAutoSize = (int)sheet.GetColumnWidth(columnIndex);
            int width = (int)(widthAfterAutoSize * 1.5);
            sheet.SetColumnWidth(columnIndex, width);
        }
    }
}
