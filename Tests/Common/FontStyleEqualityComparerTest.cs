using AwesomeExcel.Core.Comparers;
using AwesomeExcel.Models;

namespace Tests.Common;

[TestClass]
public class FontStyleEqualityComparerTest
{
    [TestMethod]
    public void Equals_Null_Null_ShouldBeTrue()
    {
        FontStyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(null, null);
        Assert.IsTrue(areEquals);
    }

    [TestMethod]
    public void Equals_NewInstance_Null_ShouldBeFalse()
    {
        FontStyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(new FontStyle(), null);
        Assert.IsFalse(areEquals);
    }

    [TestMethod]
    public void Equals_Null_NewInstance_ShouldBeFalse()
    {
        FontStyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(null, new FontStyle());
        Assert.IsFalse(areEquals);
    }

    [TestMethod]
    public void Equals_NewInstance_NewInstance_ShouldBeTrue()
    {
        FontStyleEqualityComparer comparer = new();
        bool areEquals = comparer.Equals(new FontStyle(), new FontStyle());
        Assert.IsTrue(areEquals);
    }

    [TestMethod]
    public void GetHashCode_NewInstance_ShouldBeEqualTo_GetHashCode_NewInstance()
    {
        FontStyleEqualityComparer comparer = new();
        int hashCode1 = comparer.GetHashCode(new FontStyle());
        int hashCode2 = comparer.GetHashCode(new FontStyle());
        Assert.AreEqual(hashCode1, hashCode2);
    }

    [TestMethod]
    public void GetHashCode_Instance_ShouldBeEqualTo_GetHashCode_AnotherInstanceWithSameValues()
    {
        FontStyleEqualityComparer comparer = new();

        var font1 = new FontStyle() { IsBold = true, Color = Color.BlueGray, HeightInPoints = 13, Name = "Arial" };
        var font2 = new FontStyle() { IsBold = true, Color = Color.BlueGray, HeightInPoints = 13, Name = "Arial" };

        int hashCode1 = comparer.GetHashCode(font1);
        int hashCode2 = comparer.GetHashCode(font2);
        Assert.AreEqual(hashCode1, hashCode2);
    }

    [TestMethod]
    public void GetHashCode_Instance_ShouldNotBeEqualTo_GetHashCode_InstanceWithDifferentValues()
    {
        FontStyleEqualityComparer comparer = new();

        var font1 = new FontStyle() { IsBold = true, Color = Color.BlueGray, HeightInPoints = 13, Name = "Arial" };
        var font2 = new FontStyle() { IsBold = true, Color = Color.BlueGray, HeightInPoints = 15, Name = "Arial" };

        int hashCode1 = comparer.GetHashCode(font1);
        int hashCode2 = comparer.GetHashCode(font2);
        Assert.AreNotEqual(hashCode1, hashCode2);
    }

    [TestMethod]
    public void Dictionary_Add_ReturnsTrue_Add_ReturnsFalse()
    {
        FontStyleEqualityComparer comparer = new();
        HashSet<FontStyle> dictionary = new(comparer);

        bool added1 = dictionary.Add(new FontStyle() { IsBold = true, Color = Color.BlueGray, HeightInPoints = 13, Name = "Arial" });
        Assert.IsTrue(added1);

        bool added2 = dictionary.Add(new FontStyle() { IsBold = true, Color = Color.BlueGray, HeightInPoints = 13, Name = "Arial" });
        Assert.IsFalse(added2);
    }
}
