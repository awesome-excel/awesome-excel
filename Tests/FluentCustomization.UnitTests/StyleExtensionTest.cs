using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization;

namespace AwesomeExcel.FluentCustomization.UnitTests;

[TestClass]
public class StyleExtensionTest
{
    [TestMethod]
    public void SetBorderBottomColor_ShouldSet_BorderBottomColor()
    {
        Style s = new();
        Assert.IsNull(s.BorderBottomColor);
        s.SetBorderBottomColor(Color.Ivory);
        Assert.AreEqual(s.BorderBottomColor, Color.Ivory);
    }

    [TestMethod]
    public void SetBorderBottomColor_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetBorderBottomColor(Color.Ivory);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderBottomColor_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderBottomColor(Color.Ivory));
    }

    [TestMethod]
    public void SetBorderLeftColor_ShouldSet_BorderLeftColor()
    {
        Style s = new();
        Assert.IsNull(s.BorderLeftColor);
        s.SetBorderLeftColor(Color.Green);
        Assert.AreEqual(s.BorderLeftColor, Color.Green);
    }

    [TestMethod]
    public void SetBorderLeftColor_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetBorderLeftColor(Color.Green);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderLeftColor_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderLeftColor(Color.Green));
    }

    [TestMethod]
    public void SetBorderRightColor_ShouldSet_BorderRightColor()
    {
        Style s = new();
        Assert.IsNull(s.BorderRightColor);
        s.SetBorderRightColor(Color.Black);
        Assert.AreEqual(s.BorderRightColor, Color.Black);
    }

    [TestMethod]
    public void SetBorderRightColor_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetBorderRightColor(Color.Black);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderRightColor_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderRightColor(Color.Black));
    }

    [TestMethod]
    public void SetBorderTopColor_ShouldSet_BorderTopColor()
    {
        Style s = new();
        Assert.IsNull(s.BorderTopColor);
        s.SetBorderTopColor(Color.Blue);
        Assert.AreEqual(s.BorderTopColor, Color.Blue);
    }

    [TestMethod]
    public void SetBorderTopColor_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetBorderTopColor(Color.Blue);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetBorderTopColor_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetBorderTopColor(Color.Blue));
    }

    [TestMethod]
    public void SetDateTimeFormat_ShouldSet_DateTimeFormat()
    {
        Style s = new();
        Assert.IsNull(s.DateTimeFormat);
        s.SetDateTimeFormat("dd/mm/yyyy");
        Assert.AreEqual(s.DateTimeFormat, "dd/mm/yyyy");
    }

    [TestMethod]
    public void SetDateTimeFormat_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetDateTimeFormat("dd/mm/yyyy");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetDateTimeFormat_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetDateTimeFormat("dd/mm/yyyy"));
    }

    [TestMethod]
    public void SetFillForegroundColor_ShouldSet_FillForegroundColor()
    {
        Style s = new();
        Assert.IsNull(s.FillForegroundColor);
        s.SetFillForegroundColor(Color.Yellow);
        Assert.AreEqual(s.FillForegroundColor, Color.Yellow);
    }

    [TestMethod]
    public void SetFillForegroundColor_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetFillForegroundColor(Color.Yellow);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFillForegroundColor_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFillForegroundColor(Color.Yellow));
    }

    [TestMethod]
    public void SetHorizontalAlignment_ShouldSet_HorizontalAlignment()
    {
        Style s = new();
        Assert.IsNull(s.HorizontalAlignment);

        s.SetHorizontalAlignment(HorizontalAlignment.Center);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.Center);

        s.SetHorizontalAlignment(HorizontalAlignment.CenterSelection);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.CenterSelection);

        s.SetHorizontalAlignment(HorizontalAlignment.Distributed);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.Distributed);

        s.SetHorizontalAlignment(HorizontalAlignment.Fill);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.Fill);

        s.SetHorizontalAlignment(HorizontalAlignment.General);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.General);

        s.SetHorizontalAlignment(HorizontalAlignment.Justify);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.Justify);

        s.SetHorizontalAlignment(HorizontalAlignment.Left);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.Left);

        s.SetHorizontalAlignment(HorizontalAlignment.Right);
        Assert.AreEqual(s.HorizontalAlignment, HorizontalAlignment.Right);
    }

    [TestMethod]
    public void SetHorizontalAlignment_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetHorizontalAlignment(HorizontalAlignment.Center);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetHorizontalAlignment_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetHorizontalAlignment(HorizontalAlignment.Center));
    }

    [TestMethod]
    public void SetVerticalAlignment_ShouldSet_VerticalAlignment()
    {
        Style s = new();
        Assert.IsNull(s.VerticalAlignment);

        s.SetVerticalAlignment(VerticalAlignment.Bottom);
        Assert.AreEqual(s.VerticalAlignment, VerticalAlignment.Bottom);

        s.SetVerticalAlignment(VerticalAlignment.Center);
        Assert.AreEqual(s.VerticalAlignment, VerticalAlignment.Center);

        s.SetVerticalAlignment(VerticalAlignment.Distributed);
        Assert.AreEqual(s.VerticalAlignment, VerticalAlignment.Distributed);

        s.SetVerticalAlignment(VerticalAlignment.Justify);
        Assert.AreEqual(s.VerticalAlignment, VerticalAlignment.Justify);

        s.SetVerticalAlignment(VerticalAlignment.None);
        Assert.AreEqual(s.VerticalAlignment, VerticalAlignment.None);

        s.SetVerticalAlignment(VerticalAlignment.Top);
        Assert.AreEqual(s.VerticalAlignment, VerticalAlignment.Top);
    }

    [TestMethod]
    public void SetVerticalAlignment_ShouldReturn_GivenStyleInstance()
    {
        Style s = new();
        Style returned = s.SetVerticalAlignment(VerticalAlignment.Bottom);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetVerticalAlignment_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetVerticalAlignment(VerticalAlignment.Justify));
    }

    [TestMethod]
    public void SetFontBold_ShouldSet_IsBold()
    {
        Style s = new()
        {
            FontStyle = new()
        };

        s.SetFontBold(true);
        Assert.IsTrue(s.FontStyle.IsBold);

        s.SetFontBold(false);
        Assert.IsFalse(s.FontStyle.IsBold);
    }

    [TestMethod]
    public void SetFontBold_ShouldReturn_GivenStyleInstance()
    {
        Style s = new()
        {
            FontStyle = new()
        };
        Style returned = s.SetFontBold(true);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontBold_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontBold(true));
    }

    [TestMethod]
    public void SetFontColor_ShouldSet_FontColor()
    {
        Style s = new()
        {
            FontStyle = new()
        };

        s.SetFontColor(Color.DarkBlue);
        Assert.AreEqual(s.FontStyle.Color, Color.DarkBlue);
    }

    [TestMethod]
    public void SetFontColor_ShouldReturn_GivenStyleInstance()
    {
        Style s = new()
        {
            FontStyle = new()
        };
        Style returned = s.SetFontColor(Color.DarkBlue);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontColor_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontColor(Color.DarkBlue));
    }

    [TestMethod]
    public void SetFontHeightInPoints_ShouldSet_HeightInPoints()
    {
        Style s = new()
        {
            FontStyle = new()
        };

        s.SetFontHeightInPoints(18);
        Assert.AreEqual((short)18, s.FontStyle.HeightInPoints);
    }

    [TestMethod]
    public void SetFontHeightInPoints_ShouldReturn_GivenStyleInstance()
    {
        Style s = new()
        {
            FontStyle = new()
        };
        Style returned = s.SetFontHeightInPoints(18);

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontHeightInPoints_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontHeightInPoints(18));
    }

    [TestMethod]
    public void SetFontName_ShouldSet_FontName()
    {
        Style s = new()
        {
            FontStyle = new()
        };

        s.SetFontName("FakeFontName");
        Assert.AreEqual(s.FontStyle.Name, "FakeFontName");
    }

    [TestMethod]
    public void SetFontName_ShouldReturn_GivenStyleInstance()
    {
        Style s = new()
        {
            FontStyle = new()
        };
        Style returned = s.SetFontName("FakeFontName");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetFontName_NullStyle_ShouldThrows_ArgumentNullException()
    {
        Style s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetFontName("FakeFontName"));
    }
}
