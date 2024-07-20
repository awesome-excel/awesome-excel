using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;


namespace AwesomeExcel.BridgeNpoi;

internal class StylesCache
{
    private readonly Dictionary<_Excel.Style, _NPOI.ICellStyle> cache = new(new Common.Comparers.StyleEqualityComparer());

    public _NPOI.ICellStyle Get(_Excel.Style excelStyle)
    {
        // This cache is necessary because there's a limit on how many Cell Formats can be used in a workbook.

        // Explanation:
        //     "This problem occurs when the workbook contains more than approximately 4,000 different combinations of cell
        //     formats in Excel 2003 or 64,000 different combinations in Excel 2007 and later versions."

        // Common errors:
        //     "Too many different cell formats."
        //     "Excel encountered an error and had to remove some formatting to avoid corrupting the workbook."
        //     "Excel found unreadable content in the file."
        //     "When you open a file, all the formatting is missing."

        // Official source:
        //     https://docs.microsoft.com/en-US/office/troubleshoot/excel/too-many-different-cell-formats-in-excel

        // Solution:
        //    Using a cache to re-use the same instance of (NPOI) ICellStyle for multiple cells/rows

        if (cache.TryGetValue(excelStyle, out _NPOI.ICellStyle npoiStyle))
            return npoiStyle;

        return null;
    }

    public void Add(_NPOI.ICellStyle npoiStyle, _Excel.Style excelStyle)
    {
        cache.Add(excelStyle, npoiStyle);
    }

    public void Clear()
    {
        cache.Clear();
    }
}
