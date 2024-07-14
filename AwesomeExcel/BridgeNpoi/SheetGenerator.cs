using NPOI.SS.UserModel;
using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class SheetGenerator
{
    private readonly NpoiFacade npoiFacade = new();
    private readonly _NPOI.IWorkbook npoiWorkbook;

    public SheetGenerator(_NPOI.IWorkbook npoiWorkbook)
    {
        this.npoiWorkbook = npoiWorkbook;
    }

    public void GenerateSheet(_Excel.Sheet excelSheet)
    {
        ISheet npoiSheet = string.IsNullOrWhiteSpace(excelSheet.Name)
            ? npoiWorkbook.CreateSheet()
            : npoiWorkbook.CreateSheet(excelSheet.Name);

        GenerateRows(npoiSheet, excelSheet);
        AutoSizeColumns(npoiSheet, excelSheet);
        ProtectSheet(npoiSheet, excelSheet);
    }

    private void GenerateRows(_NPOI.ISheet npoiSheet, _Excel.Sheet excelSheet)
    {
        using StyleConverterWithCache styleConverter = new(npoiSheet);
        RowsGenerator rg = new(npoiSheet, excelSheet, styleConverter);
        rg.GenerateHeaderRow();
        rg.GenerateRows();
    }

    private void ProtectSheet(_NPOI.ISheet npoiSheet, _Excel.Sheet excelSheet)
    {
        if (excelSheet.IsReadOnly)
        {
            npoiSheet.ProtectSheet("");
        }
    }

    private void AutoSizeColumns(_NPOI.ISheet npoiSheet, _Excel.Sheet excelSheet)
    {
        npoiFacade.AutoSizeColumns(npoiSheet, excelSheet.Columns?.Count ?? 0);
    }
}

