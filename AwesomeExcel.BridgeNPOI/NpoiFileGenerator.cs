﻿using AwesomeExcel.Core.Services;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

public class NpoiFileGenerator : FileGenerator<_NPOI.IWorkbook>
{
    protected override _NPOI.IWorkbook Convert(AwesomeExcel.Models.Workbook workbook)
    {
        return new WorkbookConverter().Convert(workbook);
    }

    protected override MemoryStream Write(_NPOI.IWorkbook workbook)
    {
        return new WorkbookWriter().Write(workbook);
    }
}
