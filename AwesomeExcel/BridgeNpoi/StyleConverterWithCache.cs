using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

internal class StyleConverterWithCache : StyleConverter, IDisposable
{
    private readonly StylesCache stylesCache;
    private readonly FontsCache fontsCache;
    private bool disposedValue;

    public StyleConverterWithCache(_NPOI.ISheet npoiSheet) : base(npoiSheet.Workbook)
    {
        var npoiWorkbook = npoiSheet.Workbook;
        stylesCache = new StylesCache(npoiWorkbook);
        fontsCache = new FontsCache(npoiWorkbook);
    }

    public override _NPOI.ICellStyle Convert(_Excel.Style style)
    {
        _NPOI.ICellStyle npoiStyle = stylesCache.Get(style);

        if (npoiStyle == null)
        {
            npoiStyle = base.Convert(style);
            stylesCache.Add(npoiStyle, style);
        }

        return npoiStyle;
    }

    public override _NPOI.IFont Convert(_Excel.FontStyle style)
    {
        _NPOI.IFont npoiStyle = fontsCache.Get(style);

        if (npoiStyle == null)
        {
            npoiStyle = base.Convert(style);
            fontsCache.Add(npoiStyle, style);
        }

        return npoiStyle;
    }

    private void Clear()
    {
        stylesCache.Clear();
        fontsCache.Clear();
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
                Clear();
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            disposedValue = true;
        }
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}