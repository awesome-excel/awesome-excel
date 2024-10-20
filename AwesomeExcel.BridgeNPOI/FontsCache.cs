using AwesomeExcel.Core.Comparers;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

internal class FontsCache
{
    private readonly Dictionary<AwesomeExcel.Models.FontStyle, _NPOI.IFont> cache = new(new FontStyleEqualityComparer());
    private readonly Dictionary<AwesomeExcel.Models.FontStyle, int> referenceCounter = new(new FontStyleEqualityComparer());

    public _NPOI.IFont Get(AwesomeExcel.Models.FontStyle fontStyle)
    {
        // Problem 1:
        //     This cache is necessary due to a limit on the number of Fonts that can be used in a workbook.
        //
        //     Explanation:
        //         "The maximum number of unique fonts in a workbook is limited to 32767. You should re-use fonts in your applications instead of creating a font for each cell."
        //
        //     Official source:
        //         https://poi.apache.org/components/spreadsheet/quick-guide.html#WorkingWithFonts
        //
        //     Solution:
        //         Use a cache to re-use the same instance of (NPOI) IFont for multiple cells/rows.
        //

        // Problem 2:
        //     It's necessary to track how many references of the same FontStyle have been used
        //     because there is a limit on how many times a single (NPOI) IFont instance can be used in a workbook.
        //
        //     Explanation:
        //         The maximum number of times that an IFont instance can be used in a workbook is limited to 24.
        //
        //     Official source:
        //         None
        //
        //     Solution:
        //         This cache tracks the number of references for each font to manage usage limits.

        (_NPOI.IFont npoiFont, int referenceCounter) = GetFromCache(fontStyle);

        if (npoiFont == null)
        {
            return null;
        }

        // One IFont instance can be used for styling up to 24 times 
        bool limitReached = referenceCounter >= 25;

        if (limitReached)
        {
            // Remove the font from cache.
            // In this way, a new font instance will be created for this font style
            Remove(fontStyle);
            return null;
        }

        return npoiFont;
    }

    private (_NPOI.IFont instance, int usageCount) GetFromCache(AwesomeExcel.Models.FontStyle fontStyle)
    {
        if (cache.TryGetValue(fontStyle, out _NPOI.IFont npoiFont))
        {
            referenceCounter[fontStyle]++;
            return (npoiFont, referenceCounter[fontStyle]);
        }

        return (null, 0);
    }

    public void Add(_NPOI.IFont npoiFont, AwesomeExcel.Models.FontStyle fontStyle)
    {
        cache.Add(fontStyle, npoiFont);
        referenceCounter.Add(fontStyle, 0);
    }

    private void Remove(AwesomeExcel.Models.FontStyle fontStyle)
    {
        cache.Remove(fontStyle);
        referenceCounter.Remove(fontStyle);
    }

    public void Clear()
    {
        cache.Clear();
        referenceCounter.Clear();
    }
}