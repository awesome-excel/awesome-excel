using AwesomeExcel;

namespace Tests.FluentCustomization;

[TestClass]
public class ColumnCustomizationTest
{
    [TestMethod]
    public void SetName_ShouldSet_Name()
    {
        ColumnCustomization s = new();

        s.SetName("FakeName");
        Assert.AreEqual(s.Name, "FakeName");
    }

    [TestMethod]
    public void SetName_ShouldReturn_GivenInstance()
    {
        ColumnCustomization s = new();
        ColumnCustomization returned = s.SetName("FakeName");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetName_Null_ShouldThrows_ArgumentNullException()
    {
        ColumnCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetName("FakeFontName"));
    }

    [TestMethod]
    public void Exclude_ShouldSet_Exclude()
    {
        ColumnCustomization s = new();
        s.Exclude();

        Assert.IsTrue(s.Excluded);
    }

    [TestMethod]
    public void Exclude_ShouldReturn_GivenInstance()
    {
        ColumnCustomization s = new();
        ColumnCustomization returned = s.Exclude();

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void Exlucde_Null_ShouldThrows_ArgumentNullException()
    {
        ColumnCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.Exclude());
    }
}
