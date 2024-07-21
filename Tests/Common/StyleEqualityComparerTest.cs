using AwesomeExcel;
using AwesomeExcel.Core.Comparers;

namespace Tests.Common;

[TestClass]
public class StyleEqualityComparerTest
{
    [TestMethod]
    public void Equals_Null_Null_ShouldBeTrue()
    {
        StyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(null, null);
        Assert.IsTrue(areEquals);
    }

    [TestMethod]
    public void Equals_NewInstance_Null_ShouldBeFalse()
    {
        StyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(new Style(), null);
        Assert.IsFalse(areEquals);
    }

    [TestMethod]
    public void Equals_Null_NewInstance_ShouldBeFalse()
    {
        StyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(null, new Style());
        Assert.IsFalse(areEquals);
    }

    [TestMethod]
    public void Equals_NewInstance_NewInstance_ShouldBeTrue()
    {
        StyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(new Style(), new Style());
        Assert.IsTrue(areEquals);
    }

    [TestMethod]
    public void GetHashCode_NewInstance_ShouldBeEqualTo_GetHashCode_NewInstance()
    {
        StyleEqualityComparer comparer = new();
        int hashCode1 = comparer.GetHashCode(new Style());
        int hashCode2 = comparer.GetHashCode(new Style());
        Assert.AreEqual(hashCode1, hashCode2);
    }

    [TestMethod]
    public void GetHashCode_Instance_ShouldBeEqualTo_GetHashCode_AnotherInstanceWithSameValues()
    {
        StyleEqualityComparer comparer = new();

        var style1 = new Style() { HorizontalAlignment = HorizontalAlignment.General };
        var style2 = new Style() { HorizontalAlignment = HorizontalAlignment.General };

        int hashCode1 = comparer.GetHashCode(style1);
        int hashCode2 = comparer.GetHashCode(style2);
        Assert.AreEqual(hashCode1, hashCode2);
    }

    [TestMethod]
    public void GetHashCode_Instance_ShouldNotBeEqualTo_GetHashCode_InstanceWithDifferentValues()
    {
        StyleEqualityComparer comparer = new();

        var style1 = new Style() { HorizontalAlignment = HorizontalAlignment.General };
        var style2 = new Style() { HorizontalAlignment = HorizontalAlignment.Left };

        int hashCode1 = comparer.GetHashCode(style1);
        int hashCode2 = comparer.GetHashCode(style2);
        Assert.AreNotEqual(hashCode1, hashCode2);
    }
}
