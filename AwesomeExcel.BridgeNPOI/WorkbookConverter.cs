using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

public class WorkbookConverter
{
    public _NPOI.IWorkbook Convert(Workbook excelWorkbook)
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

        foreach (Sheet excelSheet in excelWorkbook.Sheets)
        {
            sheetGenerator.GenerateSheet(excelSheet);
        }

        return npoiWorkbook;
    }

    private _NPOI.IWorkbook GetWorkbook(FileType fileType)
    {
        _NPOI.IWorkbook workbook = fileType switch
        {
            FileType.Xls => new NPOI.HSSF.UserModel.HSSFWorkbook(),
            //_Excel.FileType.Xlsx => new NPOI.XSSF.UserModel.XSSFWorkbook(),
            FileType.Xlsx => new NPOI.XSSF.Streaming.SXSSFWorkbook(),

            _ => throw new NotSupportedException(),
        };

        return workbook;
    }
}