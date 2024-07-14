using AwesomeExcel.Generator;
using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

public class NpoiFileGenerator : FileGenerator<_NPOI.IWorkbook>
{
    protected override _NPOI.IWorkbook Convert(_Excel.Workbook workbook)
    {
        return new WorkbookConverter().Convert(workbook);
    }

    protected override MemoryStream Write(_NPOI.IWorkbook workbook)
    {
        return new WorkbookWriter().Write(workbook);
    }
}
