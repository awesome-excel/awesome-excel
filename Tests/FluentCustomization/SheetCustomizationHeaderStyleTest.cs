using AwesomeExcel;

namespace Tests.FluentCustomization;

[TestClass]
public class SheetCustomizationHeaderStyleTest
{
    [TestMethod]
    public void SetBorderBottomColor_ShouldSet_BorderBottomColor()
    {
        SheetCustomization s = new();
        s.SetHeaderBorderBottomColor(Color.Ivory);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.BorderBottomColor, Color.Ivory);
    }

    [TestMethod]
    public void SetBorderBottomColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderBorderBottomColor(Color.Ivory);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderBottomColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderBorderBottomColor(Color.Ivory));
    }

    [TestMethod]
    public void SetBorderLeftColor_ShouldSet_BorderLeftColor()
    {
        SheetCustomization s = new();
        s.SetHeaderBorderLeftColor(Color.Green);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.BorderLeftColor, Color.Green);
    }

    [TestMethod]
    public void SetBorderLeftColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderBorderLeftColor(Color.Green);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderLeftColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderBorderLeftColor(Color.Green));
    }

    [TestMethod]
    public void SetBorderRightColor_ShouldSet_BorderRightColor()
    {
        SheetCustomization s = new();
        s.SetHeaderBorderRightColor(Color.Black);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.BorderRightColor, Color.Black);
    }

    [TestMethod]
    public void SetBorderRightColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderBorderRightColor(Color.Black);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderRightColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderBorderRightColor(Color.Black));
    }

    [TestMethod]
    public void SetBorderTopColor_ShouldSet_BorderTopColor()
    {
        SheetCustomization s = new();
        s.SetHeaderBorderTopColor(Color.Blue);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.BorderTopColor, Color.Blue);
    }

    [TestMethod]
    public void SetBorderTopColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderBorderTopColor(Color.Blue);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderTopColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderBorderTopColor(Color.Blue));
    }

    [TestMethod]
    public void SetDateTimeFormat_ShouldSet_DateTimeFormat()
    {
        SheetCustomization s = new();
        s.SetHeaderDateTimeFormat("dd/mm/yyyy");

        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.DateTimeFormat, "dd/mm/yyyy");
    }

    [TestMethod]
    public void SetDateTimeFormat_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderDateTimeFormat("dd/mm/yyyy");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetDateTimeFormat_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderDateTimeFormat("dd/mm/yyyy"));
    }

    [TestMethod]
    public void SetFillForegroundColor_ShouldSet_FillForegroundColor()
    {
        SheetCustomization s = new();
        s.SetHeaderFillForegroundColor(Color.Yellow);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.FillForegroundColor, Color.Yellow);
    }

    [TestMethod]
    public void SetFillForegroundColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderFillForegroundColor(Color.Yellow);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFillForegroundColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderFillForegroundColor(Color.Yellow));
    }

    [TestMethod]
    public void SetHorizontalAlignment_ShouldSet_HorizontalAlignment()
    {
        SheetCustomization s = new();

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.Center);
        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Center);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.CenterSelection);
        Assert.AreEqual(actual: s.HeaderStyle.HorizontalAlignment, expected: HorizontalAlignment.CenterSelection);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.Distributed);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Distributed);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.Fill);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Fill);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.General);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.General);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.Justify);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Justify);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.Left);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Left);

        s.SetHeaderHorizontalAlignment(HorizontalAlignment.Right);
        Assert.AreEqual(s.HeaderStyle.HorizontalAlignment, HorizontalAlignment.Right);
    }

    [TestMethod]
    public void SetHorizontalAlignment_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderHorizontalAlignment(HorizontalAlignment.Center);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetHorizontalAlignment_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderHorizontalAlignment(HorizontalAlignment.Center));
    }

    [TestMethod]
    public void SetVerticalAlignment_ShouldSet_VerticalAlignment()
    {
        SheetCustomization s = new();

        s.SetHeaderVerticalAlignment(VerticalAlignment.Bottom);
        Assert.IsNotNull(s.HeaderStyle);
        Assert.AreEqual(s.HeaderStyle.VerticalAlignment, VerticalAlignment.Bottom);

        s.SetHeaderVerticalAlignment(VerticalAlignment.Center);
        Assert.AreEqual(s.HeaderStyle.VerticalAlignment, VerticalAlignment.Center);

        s.SetHeaderVerticalAlignment(VerticalAlignment.Distributed);
        Assert.AreEqual(s.HeaderStyle.VerticalAlignment, VerticalAlignment.Distributed);

        s.SetHeaderVerticalAlignment(VerticalAlignment.Justify);
        Assert.AreEqual(s.HeaderStyle.VerticalAlignment, VerticalAlignment.Justify);

        s.SetHeaderVerticalAlignment(VerticalAlignment.None);
        Assert.AreEqual(s.HeaderStyle.VerticalAlignment, VerticalAlignment.None);

        s.SetHeaderVerticalAlignment(VerticalAlignment.Top);
        Assert.AreEqual(s.HeaderStyle.VerticalAlignment, VerticalAlignment.Top);
    }

    [TestMethod]
    public void SetVerticalAlignment_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderVerticalAlignment(VerticalAlignment.Bottom);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetVerticalAlignment_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderVerticalAlignment(VerticalAlignment.Justify));
    }

    [TestMethod]
    public void SetFontBold_ShouldSet_IsBold()
    {
        SheetCustomization s = new();

        s.SetHeaderFontBold(true);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.IsNotNull(s.HeaderStyle.FontStyle);
        Assert.IsTrue(s.HeaderStyle.FontStyle.IsBold);

        s.SetHeaderFontBold(false);
        Assert.IsFalse(s.HeaderStyle.FontStyle.IsBold);
    }

    [TestMethod]
    public void SetFontBold_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderFontBold(true);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontBold_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderFontBold(true));
    }

    [TestMethod]
    public void SetFontColor_ShouldSet_FontColor()
    {
        SheetCustomization s = new();

        s.SetHeaderFontColor(Color.DarkBlue);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.IsNotNull(s.HeaderStyle.FontStyle);
        Assert.AreEqual(s.HeaderStyle.FontStyle.Color, Color.DarkBlue);
    }

    [TestMethod]
    public void SetFontColor_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderFontColor(Color.DarkBlue);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontColor_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderFontColor(Color.DarkBlue));
    }

    [TestMethod]
    public void SetFontHeightInPoints_ShouldSet_HeightInPoints()
    {
        SheetCustomization s = new();

        s.SetHeaderFontHeightInPoints(18);

        Assert.IsNotNull(s.HeaderStyle);
        Assert.IsNotNull(s.HeaderStyle.FontStyle);
        Assert.AreEqual((short)18, s.HeaderStyle.FontStyle.HeightInPoints);
    }

    [TestMethod]
    public void SetFontHeightInPoints_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderFontHeightInPoints(18);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontHeightInPoints_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderFontHeightInPoints(18));
    }

    [TestMethod]
    public void SetFontName_ShouldSet_FontName()
    {
        SheetCustomization s = new();

        s.SetHeaderFontName("FakeFontName");

        Assert.IsNotNull(s.HeaderStyle);
        Assert.IsNotNull(s.HeaderStyle.FontStyle);
        Assert.AreEqual(s.HeaderStyle.FontStyle.Name, "FakeFontName");
    }

    [TestMethod]
    public void SetFontName_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetHeaderFontName("FakeFontName");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontName_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHeaderFontName("FakeFontName"));
    }
}
