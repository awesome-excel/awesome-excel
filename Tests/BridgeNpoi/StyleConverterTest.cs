using AwesomeExcel.BridgeNpoi;
using AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace Tests.BridgeNpoi;

[TestClass]
public class StyleConverterTest
{
    private _NPOI.IWorkbook GetNpoiWorkbookXlsx()
    {
        return new NPOI.XSSF.UserModel.XSSFWorkbook();
    }
    private _NPOI.IWorkbook GetNpoiWorkbookXls()
    {
        return new NPOI.HSSF.UserModel.HSSFWorkbook();
    }


    [TestMethod]
    public void Convert_NullWorkbook_ShouldThrow_ArgumentNullException()
    {
        Assert.ThrowsException<ArgumentNullException>(() => new StyleConverter(null));
    }

    [TestMethod]
    public void Convert_NullStyle_Returns_NullNPOIStyle()
    {
        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXlsx();
        StyleConverter converter = new StyleConverter(npoiWorkbook);
        Style s = null;
        _NPOI.ICellStyle npoiStyle = converter.Convert(s);

        Assert.AreEqual(null, npoiStyle);
    }

    [TestMethod]
    public void Convert_NullFontStyle_Returns_NullNPOIFontStyle()
    {
        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXlsx();
        StyleConverter converter = new StyleConverter(npoiWorkbook);
        FontStyle fs = null;
        _NPOI.IFont npoiStyle = converter.Convert(fs);

        Assert.AreEqual(null, npoiStyle);
    }

    [TestMethod]
    public void Convert_Style_BorderTopColor_Returns_BorderTopColor()
    {
        Style blueGrayBorderTopColor = new()
        {
            BorderTopColor = Color.BlueGray
        };

        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXlsx();
        StyleConverter converter = new(npoiWorkbook);
        _NPOI.ICellStyle npoiStyle = converter.Convert(blueGrayBorderTopColor);

        short expected = npoiStyle.TopBorderColor;
        short actual = (short)blueGrayBorderTopColor.BorderTopColor;
        Assert.AreEqual(expected, actual);
    }

    [TestMethod]
    public void Convert_Style_GenericStyle_Xlsx_Returns_Style()
    {
        Style excelStyle = new()
        {
            BorderTopColor = Color.BlueGray,
            BorderBottomColor = Color.DarkTeal,
            BorderLeftColor = Color.Indigo,
            BorderRightColor = Color.Pink,
            DateTimeFormat = "yyyy/mm/dd hh:mm",
            FillForegroundColor = Color.Yellow,
            FillPattern = FillPattern.LeastDots,
            FontStyle = null,
            HorizontalAlignment = HorizontalAlignment.Distributed,
            VerticalAlignment = VerticalAlignment.Top
        };

        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXlsx();
        StyleConverter converter = new(npoiWorkbook);
        _NPOI.ICellStyle npoiStyle = converter.Convert(excelStyle);

        Assert.AreEqual(expected: (short)excelStyle.BorderTopColor, actual: npoiStyle.TopBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.BorderBottomColor, actual: npoiStyle.BottomBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.BorderLeftColor, actual: npoiStyle.LeftBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.BorderRightColor, actual: npoiStyle.RightBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.FillForegroundColor, actual: npoiStyle.FillForegroundColor);
        Assert.AreEqual(expected: (short)excelStyle.FillPattern, actual: (short)npoiStyle.FillPattern);
        Assert.AreEqual(expected: excelStyle.HorizontalAlignment.ToString(), actual: npoiStyle.Alignment.ToString());
        Assert.AreEqual(expected: excelStyle.VerticalAlignment.ToString(), actual: npoiStyle.VerticalAlignment.ToString());

        short expectedDataFormat = npoiWorkbook.CreateDataFormat().GetFormat(excelStyle.DateTimeFormat);
        short actualDataFormat = npoiStyle.DataFormat;
        Assert.AreEqual(expected: expectedDataFormat, actual: actualDataFormat);
    }


    [TestMethod]
    public void Convert_Style_GenericStyle_Xls_Returns_Style()
    {
        Style excelStyle = new()
        {
            BorderTopColor = Color.BlueGray,
            BorderBottomColor = Color.DarkTeal,
            BorderLeftColor = Color.Indigo,
            BorderRightColor = Color.Pink,
            DateTimeFormat = "yyyy/mm/dd hh:mm",
            FillForegroundColor = Color.Yellow,
            FillPattern = FillPattern.LeastDots,
            FontStyle = null,
            HorizontalAlignment = HorizontalAlignment.Distributed,
            VerticalAlignment = VerticalAlignment.Top
        };

        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXls();
        StyleConverter converter = new(npoiWorkbook);
        _NPOI.ICellStyle npoiStyle = converter.Convert(excelStyle);

        Assert.AreEqual(expected: (short)excelStyle.BorderTopColor, actual: npoiStyle.TopBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.BorderBottomColor, actual: npoiStyle.BottomBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.BorderLeftColor, actual: npoiStyle.LeftBorderColor);
        Assert.AreEqual(expected: (short)excelStyle.BorderRightColor, actual: npoiStyle.RightBorderColor);

        Assert.AreEqual(expected: (short)excelStyle.FillForegroundColor, actual: npoiStyle.FillForegroundColor);
        Assert.AreEqual(expected: (short)excelStyle.FillPattern, actual: (short)npoiStyle.FillPattern);
        //Assert.AreEqual(expected: (short)excelStyle.FontStyle, actual: null);
        Assert.AreEqual(expected: excelStyle.HorizontalAlignment.ToString(), actual: npoiStyle.Alignment.ToString());
        Assert.AreEqual(expected: excelStyle.VerticalAlignment.ToString(), actual: npoiStyle.VerticalAlignment.ToString());

        short expectedDataFormat = npoiWorkbook.CreateDataFormat().GetFormat(excelStyle.DateTimeFormat);
        Assert.AreEqual(expected: expectedDataFormat, actual: npoiStyle.DataFormat);
    }

    [TestMethod]
    public void Convert_FontStyle_Color_Returns_FontStyle()
    {
        FontStyle excelFont = new()
        {
            Color = Color.Blue
        };

        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXlsx();
        StyleConverter converter = new(npoiWorkbook);
        _NPOI.IFont npoiFont = converter.Convert(excelFont);

        short actual = npoiFont.Color;
        short expected = (short)excelFont.Color;
        Assert.AreEqual(expected: expected, actual: actual);
    }

    [TestMethod]
    public void Convert_FontStyle_Generic_Returns_FontStyle()
    {
        FontStyle excelFont = new()
        {
            Color = Color.Blue,
            HeightInPoints = 13,
            IsBold = true,
            Name = "Arial"
        };

        _NPOI.IWorkbook npoiWorkbook = GetNpoiWorkbookXlsx();
        StyleConverter converter = new(npoiWorkbook);
        _NPOI.IFont npoiFont = converter.Convert(excelFont);

        Assert.AreEqual(expected: (short)excelFont.Color, actual: npoiFont.Color);
        Assert.AreEqual(expected: (double)excelFont.HeightInPoints, actual: npoiFont.FontHeightInPoints);
        Assert.AreEqual(expected: excelFont.IsBold, actual: npoiFont.IsBold);
        Assert.AreEqual(expected: excelFont.Name, actual: npoiFont.FontName);
    }
}
