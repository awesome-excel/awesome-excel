using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AwesomeExcel.BridgeNPOI;

public class WorkbookWriter
{
    public MemoryStream Write(IWorkbook workbook)
    {
        if (workbook is null)
        {
            throw new ArgumentNullException(nameof(workbook));
        }

        if (workbook is XSSFWorkbook xssfWorkbook)
        {
            MemoryStream ms = new();
            xssfWorkbook.Write(ms, leaveOpen: true);
            ms.Seek(0, SeekOrigin.Begin);
            return ms;
        }
        else
        {
            var ms = new NpoiMemoryStream
            {
                AllowClose = false
            };
            workbook.Write(ms);
            ms.Flush();
            ms.Seek(0, SeekOrigin.Begin);
            ms.AllowClose = true;
            return ms;
        }
    }

    private class NpoiMemoryStream : MemoryStream
    {
        // MemoryStream seems be closed after NPOI workbook.write?
        // https://stackoverflow.com/questions/22931582/memorystream-seems-be-closed-after-npoi-workbook-write

        public NpoiMemoryStream()
        {
            // We always want to close streams by default to
            // force the developer to make the conscious decision
            // to disable it.  Then, they're more apt to remember
            // to re-enable it.  The last thing you want is to
            // enable memory leaks by default.  ;-)
            AllowClose = true;
        }

        public bool AllowClose { get; set; }

        public override void Close()
        {
            if (AllowClose)
                base.Close();
        }
    }
}
