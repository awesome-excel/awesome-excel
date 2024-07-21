using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

internal class StyleConverterWithCache : StyleConverter, IDisposable
{
    private readonly StylesCache stylesCache;
    private readonly FontsCache fontsCache;
    private bool disposedValue;

    public StyleConverterWithCache(_NPOI.ISheet npoiSheet) : base(npoiSheet.Workbook)
    {
        var npoiWorkbook = npoiSheet.Workbook;
        stylesCache = new StylesCache();
        fontsCache = new FontsCache();
    }

    public override _NPOI.ICellStyle? Convert(Style? style)
    {
        if (style is null)
        {
            return null;
        }

        _NPOI.ICellStyle npoiStyle = stylesCache.Get(style);

        if (npoiStyle == null)
        {
            npoiStyle = base.Convert(style);
            stylesCache.Add(npoiStyle, style);
        }

        return npoiStyle;
    }

    public override _NPOI.IFont? Convert(FontStyle? style)
    {
        if (style is null)
        {
            return null;
        }

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