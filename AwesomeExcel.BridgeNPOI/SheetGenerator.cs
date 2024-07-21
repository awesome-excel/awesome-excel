using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

internal class SheetGenerator
{
    private readonly _NPOI.IWorkbook npoiWorkbook;

    public SheetGenerator(_NPOI.IWorkbook npoiWorkbook)
    {
        this.npoiWorkbook = npoiWorkbook;
    }

    public void GenerateSheet(Sheet excelSheet)
    {
        _NPOI.ISheet npoiSheet = string.IsNullOrWhiteSpace(excelSheet.Name)
            ? npoiWorkbook.CreateSheet()
            : npoiWorkbook.CreateSheet(excelSheet.Name);

        GenerateRows(npoiSheet, excelSheet);
        AutoSizeColumns(npoiSheet, excelSheet);
        ProtectSheet(npoiSheet, excelSheet);
    }

    private void GenerateRows(_NPOI.ISheet npoiSheet, Sheet excelSheet)
    {
        using StyleConverterWithCache styleConverter = new(npoiSheet);
        RowsGenerator rg = new(npoiSheet, excelSheet, styleConverter);
        rg.GenerateHeaderRow();
        rg.GenerateRows();
    }

    private void ProtectSheet(_NPOI.ISheet npoiSheet, Sheet excelSheet)
    {
        if (excelSheet.IsReadOnly)
        {
            npoiSheet.ProtectSheet("");
        }
    }

    private void AutoSizeColumns(_NPOI.ISheet npoiSheet, Sheet excelSheet)
    {
        // NPOI AutoSizeColumns is too slow 

        int columnsCount = excelSheet.Columns?.Count ?? 0;

        for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
        {
            const int maxAllowedWidth = 255 * 256;
            const int characters = 15; // arbitrary
            const int standardMinimumWidth = (int)(characters * 1.14388) * 256;
            int currentWidth = (int)npoiSheet.GetColumnWidth(columnIndex);
            int columnMinimumWidth = getColumnMinimumWidth(excelSheet, columnIndex);

            int newWidth = standardMinimumWidth;
            newWidth = Math.Max(newWidth, currentWidth);
            newWidth = Math.Max(newWidth, columnMinimumWidth);
            newWidth = Math.Min(newWidth, maxAllowedWidth);

            npoiSheet.SetColumnWidth(columnIndex, newWidth);
        }

        string getDateTimeFormat(Sheet excelSheet, int columnIndex)
        {
            return excelSheet.Rows[0].Cells[columnIndex].Style?.DateTimeFormat ?? excelSheet.Columns[columnIndex].Style?.DateTimeFormat;
        }

        string getTodayExcelFormat(string _dateTimeFormat)
        {
            string? excelFormat = _dateTimeFormat?.Replace("mmmm", "MMMMM")?.Replace("YYYY", "yyyy");
            return DateTime.Now.ToString(excelFormat);
        }

        int getDateTimeMinimumWidth(string _dateTimeFormat)
        {
            string _today = getTodayExcelFormat(_dateTimeFormat);
            int maxNumCharacters = _today.Length;
            int columnMinimumWidth = (int)(maxNumCharacters * 1.14388) * 256;
            columnMinimumWidth = (int)(columnMinimumWidth * 1.6); // arbitrary multiplier for datetime
            return columnMinimumWidth;
        }

        int getStringMinimumWidth(int length)
        {
            int columnMinimumWidth = (int)(length * 1.14388) * 256;
            columnMinimumWidth = (int)(columnMinimumWidth * 1.3); // arbitrary multiplier for string
            return columnMinimumWidth;
        }

        int getStringCellLength(int rowIndex, int columnIndex)
        {
            Row row = excelSheet.Rows[rowIndex];
            Cell cell = row.Cells[columnIndex];
            string str = (string)cell.Value;
            return str?.Length ?? 0;
        }

        int getLongestStringCharactersCount(Sheet excelSheet, int columnIndex)
        {
            int longestString = 0;

            for (int rowIndex = 0; rowIndex < excelSheet.Rows.Count && rowIndex < 100; rowIndex++)
            {
                int length = getStringCellLength(rowIndex, columnIndex);

                if (length > longestString)
                {
                    longestString = length;
                }
            }

            return longestString;
        }

        int getColumnMinimumWidth(Sheet excelSheet, int columnIndex)
        {
            int columnMinimumWidth = 0;

            Column _column = excelSheet.Columns[columnIndex];

            if (_column.ColumnType == ColumnType.DateTime)
            {
                string _dateTimeFormat = getDateTimeFormat(excelSheet, columnIndex);
                columnMinimumWidth = getDateTimeMinimumWidth(_dateTimeFormat);
            }
            else if (_column.ColumnType == ColumnType.String)
            {
                int count = getLongestStringCharactersCount(excelSheet, columnIndex);
                columnMinimumWidth = getStringMinimumWidth(count);
            }

            return columnMinimumWidth;
        }
    }
}

