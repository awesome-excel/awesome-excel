using AwesomeExcel;

namespace Tests.FluentCustomization;

[TestClass]
public class SheetCustomizationStyleTest
{
    [TestMethod]
    public void SetBorderBottomColor_ShouldSet_BorderBottomColor()
    {
        SheetCustomization s = new();
        s.SetBorderBottomColor(Color.Ivory);

        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.BorderBottomColor, Color.Ivory);
    }

    [TestMethod]
    public void SetBorderBottomColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetBorderBottomColor(Color.Ivory);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderBottomColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderBottomColor(Color.Ivory));
    }

    [TestMethod]
    public void SetBorderLeftColor_ShouldSet_BorderLeftColor()
    {
        SheetCustomization s = new();
        s.SetBorderLeftColor(Color.Green);

        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.BorderLeftColor, Color.Green);
    }

    [TestMethod]
    public void SetBorderLeftColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetBorderLeftColor(Color.Green);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderLeftColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderLeftColor(Color.Green));
    }

    [TestMethod]
    public void SetBorderRightColor_ShouldSet_BorderRightColor()
    {
        SheetCustomization s = new();
        s.SetBorderRightColor(Color.Black);

        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.BorderRightColor, Color.Black);
    }

    [TestMethod]
    public void SetBorderRightColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetBorderRightColor(Color.Black);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderRightColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderRightColor(Color.Black));
    }

    [TestMethod]
    public void SetBorderTopColor_ShouldSet_BorderTopColor()
    {
        SheetCustomization s = new();
        s.SetBorderTopColor(Color.Blue);

        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.BorderTopColor, Color.Blue);
    }

    [TestMethod]
    public void SetBorderTopColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetBorderTopColor(Color.Blue);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderTopColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderTopColor(Color.Blue));
    }

    [TestMethod]
    public void SetDateTimeFormat_ShouldSet_DateTimeFormat()
    {
        SheetCustomization s = new();
        s.SetDateTimeFormat("dd/mm/yyyy");

        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.DateTimeFormat, "dd/mm/yyyy");
    }

    [TestMethod]
    public void SetDateTimeFormat_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetDateTimeFormat("dd/mm/yyyy");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetDateTimeFormat_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetDateTimeFormat("dd/mm/yyyy"));
    }

    [TestMethod]
    public void SetFillForegroundColor_ShouldSet_FillForegroundColor()
    {
        SheetCustomization s = new();
        s.SetFillForegroundColor(Color.Yellow);

        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.FillForegroundColor, Color.Yellow);
    }

    [TestMethod]
    public void SetFillForegroundColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetFillForegroundColor(Color.Yellow);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFillForegroundColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFillForegroundColor(Color.Yellow));
    }

    [TestMethod]
    public void SetHorizontalAlignment_ShouldSet_HorizontalAlignment()
    {
        SheetCustomization s = new();

        s.SetHorizontalAlignment(HorizontalAlignment.Center);
        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.Center);

        s.SetHorizontalAlignment(HorizontalAlignment.CenterSelection);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.CenterSelection);

        s.SetHorizontalAlignment(HorizontalAlignment.Distributed);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.Distributed);

        s.SetHorizontalAlignment(HorizontalAlignment.Fill);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.Fill);

        s.SetHorizontalAlignment(HorizontalAlignment.General);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.General);

        s.SetHorizontalAlignment(HorizontalAlignment.Justify);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.Justify);

        s.SetHorizontalAlignment(HorizontalAlignment.Left);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.Left);

        s.SetHorizontalAlignment(HorizontalAlignment.Right);
        Assert.AreEqual(s.Style.HorizontalAlignment, HorizontalAlignment.Right);
    }

    [TestMethod]
    public void SetHorizontalAlignment_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHorizontalAlignment(HorizontalAlignment.Center);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetHorizontalAlignment_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHorizontalAlignment(HorizontalAlignment.Center));
    }

    [TestMethod]
    public void SetVerticalAlignment_ShouldSet_VerticalAlignment()
    {
        SheetCustomization s = new();

        s.SetVerticalAlignment(VerticalAlignment.Bottom);
        Assert.IsNotNull(s.Style);
        Assert.AreEqual(s.Style.VerticalAlignment, VerticalAlignment.Bottom);

        s.SetVerticalAlignment(VerticalAlignment.Center);
        Assert.AreEqual(s.Style.VerticalAlignment, VerticalAlignment.Center);

        s.SetVerticalAlignment(VerticalAlignment.Distributed);
        Assert.AreEqual(s.Style.VerticalAlignment, VerticalAlignment.Distributed);

        s.SetVerticalAlignment(VerticalAlignment.Justify);
        Assert.AreEqual(s.Style.VerticalAlignment, VerticalAlignment.Justify);

        s.SetVerticalAlignment(VerticalAlignment.None);
        Assert.AreEqual(s.Style.VerticalAlignment, VerticalAlignment.None);

        s.SetVerticalAlignment(VerticalAlignment.Top);
        Assert.AreEqual(s.Style.VerticalAlignment, VerticalAlignment.Top);
    }

    [TestMethod]
    public void SetVerticalAlignment_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetVerticalAlignment(VerticalAlignment.Bottom);

        Assert.IsNotNull(s.Style);
        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetVerticalAlignment_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetVerticalAlignment(VerticalAlignment.Justify));
    }

    [TestMethod]
    public void SetFontBold_ShouldSet_IsBold()
    {
        SheetCustomization s = new();

        s.SetFontBold(true);

        Assert.IsNotNull(s.Style);
        Assert.IsNotNull(s.Style.FontStyle);
        Assert.IsTrue(s.Style.FontStyle.IsBold);

        s.SetFontBold(false);
        Assert.IsFalse(s.Style.FontStyle.IsBold);
    }

    [TestMethod]
    public void SetFontBold_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetFontBold(true);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontBold_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontBold(true));
    }

    [TestMethod]
    public void SetFontColor_ShouldSet_FontColor()
    {
        SheetCustomization s = new();

        s.SetFontColor(Color.DarkBlue);

        Assert.IsNotNull(s.Style);
        Assert.IsNotNull(s.Style.FontStyle);
        Assert.AreEqual(s.Style.FontStyle.Color, Color.DarkBlue);
    }

    [TestMethod]
    public void SetFontColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetFontColor(Color.DarkBlue);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontColor(Color.DarkBlue));
    }

    [TestMethod]
    public void SetFontHeightInPoints_ShouldSet_HeightInPoints()
    {
        SheetCustomization s = new();

        s.SetFontHeightInPoints(18);

        Assert.IsNotNull(s.Style);
        Assert.IsNotNull(s.Style.FontStyle);
        Assert.AreEqual((short)18, s.Style.FontStyle.HeightInPoints);
    }

    [TestMethod]
    public void SetFontHeightInPoints_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetFontHeightInPoints(18);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontHeightInPoints_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontHeightInPoints(18));
    }

    [TestMethod]
    public void SetFontName_ShouldSet_FontName()
    {
        SheetCustomization s = new();

        s.SetFontName("FakeFontName");

        Assert.IsNotNull(s.Style);
        Assert.IsNotNull(s.Style.FontStyle);
        Assert.AreEqual(s.Style.FontStyle.Name, "FakeFontName");
    }

    [TestMethod]
    public void SetFontName_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetFontName("FakeFontName");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontName_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontName("FakeFontName"));
    }
}
