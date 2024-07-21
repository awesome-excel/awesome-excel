using AwesomeExcel;
using AwesomeExcel.Core.Comparers;
using AwesomeExcel.Core.Services;

namespace Tests.Common;

[TestClass]
public class StylesMergerTest
{
    [TestMethod]
    public void Merge_Null_Returns_Null()
    {
        Style style = null;

        StylesMerger merger = new();
        Style result = merger.Merge(style);

        Assert.IsNull(result);
    }

    [TestMethod]
    public void Merge_Null_Null_Returns_Null()
    {
        Style style1 = null;
        Style style2 = null;

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Assert.IsNull(result);
    }

    [TestMethod]
    public void Merge_Instance_Returns_DifferentInstance()
    {
        Style style = new();

        StylesMerger merger = new();
        Style result = merger.Merge(style);

        Assert.IsNotNull(result);
        Assert.AreNotEqual(style, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(style, result);
        Assert.IsTrue(areEqualsByComparison);
    }

    [TestMethod]
    public void Merge_Instance_Null_Returns_DifferentInstance()
    {
        Style style1 = new();
        Style style2 = null;

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(style1, result);
        Assert.IsTrue(areEqualsByComparison);
    }

    [TestMethod]
    public void Merge_Null_Instance_Returns_DifferentInstance()
    {
        Style style1 = null;
        Style style2 = new();

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(style2, result);
        Assert.IsTrue(areEqualsByComparison);
    }

    [TestMethod]
    public void Merge_Style_Null_Returns_DifferentInstanceWithSameValues()
    {
        Style style1 = new()
        {
            BorderBottomColor = Color.Aqua,
            BorderLeftColor = Color.Blue,
            BorderRightColor = Color.BrightGreen,
            BorderTopColor = Color.DarkGreen,
            DateTimeFormat = "yyyy/mm/dd",
            FillForegroundColor = Color.DarkYellow,
            FillPattern = FillPattern.LeastDots,
            HorizontalAlignment = HorizontalAlignment.Fill,
            VerticalAlignment = VerticalAlignment.Top,
            // FontStyle
        };
        Style style2 = null;

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(style1, result);
        Assert.IsTrue(areEqualsByComparison);
    }

    [TestMethod]
    public void Merge_TwoStylesWithSameValues_Returns_NewStyleWithSameValues()
    {
        Style style1 = new()
        {
            BorderBottomColor = Color.Aqua,
            BorderLeftColor = Color.Blue,
            BorderRightColor = Color.BrightGreen,
            BorderTopColor = Color.DarkGreen,
            DateTimeFormat = "yyyy/mm/dd",
            FillForegroundColor = Color.DarkYellow,
            FillPattern = FillPattern.LeastDots,
            HorizontalAlignment = HorizontalAlignment.Fill,
            VerticalAlignment = VerticalAlignment.Top,
            // FontStyle
        };
        Style style2 = new()
        {
            BorderBottomColor = Color.Aqua,
            BorderLeftColor = Color.Blue,
            BorderRightColor = Color.BrightGreen,
            BorderTopColor = Color.DarkGreen,
            DateTimeFormat = "yyyy/mm/dd",
            FillForegroundColor = Color.DarkYellow,
            FillPattern = FillPattern.LeastDots,
            HorizontalAlignment = HorizontalAlignment.Fill,
            VerticalAlignment = VerticalAlignment.Top,
            // FontStyle
        };

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(style1, result);
        Assert.IsTrue(areEqualsByComparison);
    }

    [TestMethod]
    public void Merge_TwoDifferentStyles_Returns_NewStyle()
    {
        Style style1 = new()
        {
            BorderBottomColor = Color.Aqua,
            BorderLeftColor = Color.Blue,
            BorderRightColor = Color.BrightGreen,
            BorderTopColor = Color.DarkGreen,
            FillForegroundColor = Color.DarkYellow,
            FillPattern = FillPattern.LeastDots,
            HorizontalAlignment = HorizontalAlignment.Fill,
            VerticalAlignment = VerticalAlignment.Top,
            // FontStyle
        };
        Style style2 = new()
        {
            BorderBottomColor = Color.Black,
            BorderLeftColor = Color.BlueGray,
            BorderRightColor = Color.Brown,
            BorderTopColor = Color.DarkBlue,
            FillForegroundColor = Color.Gray_50_Percent,
            FillPattern = FillPattern.Bricks,
            HorizontalAlignment = HorizontalAlignment.CenterSelection,
            VerticalAlignment = VerticalAlignment.Distributed
        };

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Style expected = new()
        {
            BorderBottomColor = Color.Black,
            BorderLeftColor = Color.BlueGray,
            BorderRightColor = Color.Brown,
            BorderTopColor = Color.DarkBlue,
            FillForegroundColor = Color.Gray_50_Percent,
            FillPattern = FillPattern.Bricks,
            HorizontalAlignment = HorizontalAlignment.CenterSelection,
            VerticalAlignment = VerticalAlignment.Distributed
        };
        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(expected, result);
        Assert.IsTrue(areEqualsByComparison);
    }


    [TestMethod]
    public void Merge_TwoDifferentStylesWithNullValues_Returns_NewStyleWithInheritedValues()
    {
        Style style1 = new()
        {
            BorderBottomColor = Color.Aqua,
            BorderLeftColor = Color.Blue,
            BorderRightColor = Color.BrightGreen,
            BorderTopColor = Color.DarkGreen,
            FillForegroundColor = Color.DarkYellow,
            FillPattern = FillPattern.LeastDots,
            HorizontalAlignment = HorizontalAlignment.Fill,
            VerticalAlignment = VerticalAlignment.Top,
            // FontStyle
        };
        Style style2 = new()
        {
            BorderBottomColor = null,
            BorderLeftColor = Color.BlueGray,
            BorderRightColor = Color.Brown,
            BorderTopColor = Color.DarkBlue,
            FillForegroundColor = null,
            FillPattern = FillPattern.Bricks,
            HorizontalAlignment = HorizontalAlignment.CenterSelection,
            VerticalAlignment = null
        };

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);

        Style expected = new()
        {
            BorderBottomColor = Color.Aqua,
            BorderLeftColor = Color.BlueGray,
            BorderRightColor = Color.Brown,
            BorderTopColor = Color.DarkBlue,
            FillForegroundColor = Color.DarkYellow,
            FillPattern = FillPattern.Bricks,
            HorizontalAlignment = HorizontalAlignment.CenterSelection,
            VerticalAlignment = VerticalAlignment.Top
        };
        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(expected, result);
        Assert.IsTrue(areEqualsByComparison);
    }

    [TestMethod]
    public void Merge_TwoStylesWithSamePropertySetWithDifferentValue_Returns_NewInstance()
    {
        Style style1 = new()
        {
            BorderBottomColor = Color.Gold
        };
        Style style2 = new()
        {
            BorderBottomColor = Color.DarkPurple
        };

        StylesMerger merger = new();
        Style result = merger.Merge(style1, style2);
        Style expected = new()
        {
            BorderBottomColor = Color.DarkPurple
        };

        Assert.IsNotNull(result);
        Assert.AreNotEqual(style1, result);
        Assert.AreNotEqual(style2, result);

        bool areEqualsByComparison = new StyleEqualityComparer().Equals(expected, result);
        Assert.IsTrue(areEqualsByComparison);
    }
}
