using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

public class WorkbookConverter
{
    public _NPOI.IWorkbook Convert(_Excel.Workbook excelWorkbook)
    {
        if (excelWorkbook is null)
        {
            throw new ArgumentNullException(nameof(excelWorkbook));
        }

        if (excelWorkbook.Sheets is null || excelWorkbook.Sheets.Count == 0)
        {
            throw new InvalidOperationException();
        }

        if (excelWorkbook.Sheets.Any(sheet => sheet is null))
        {
            throw new InvalidOperationException();
        }

        _NPOI.IWorkbook npoiWorkbook = GetWorkbook(excelWorkbook.FileType);
        SheetGenerator sheetGenerator = new(npoiWorkbook);

        foreach (_Excel.Sheet excelSheet in excelWorkbook.Sheets)
        {
            sheetGenerator.GenerateSheet(excelSheet);
        }

        return npoiWorkbook;
    }

    private _NPOI.IWorkbook GetWorkbook(_Excel.FileType fileType)
    {
        _NPOI.IWorkbook workbook = fileType switch
        {
            _Excel.FileType.Xls => new NPOI.HSSF.UserModel.HSSFWorkbook(),
            _Excel.FileType.Xlsx => new NPOI.XSSF.UserModel.XSSFWorkbook(),

            _ => throw new NotSupportedException(),
        };

        return workbook;
    }
}