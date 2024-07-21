using AwesomeExcel;

namespace Tests.FluentCustomization;

[TestClass]
public class SheetCustomizationTest
{
    [TestMethod]
    public void SetName_ShouldSet_Name()
    {
        SheetCustomization s = new();

        s.SetName("FakeName");
        Assert.AreEqual(s.Name, "FakeName");
    }

    [TestMethod]
    public void SetName_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.SetName("FakeName");

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void SetName_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.SetName("FakeFontName"));
    }

    [TestMethod]
    public void Protect_ShouldSet_Exclude()
    {
        SheetCustomization s = new();
        s.Protect();

        Assert.IsTrue(s.IsReadOnly);
    }

    [TestMethod]
    public void Protect_ShouldReturn_GivenInstance()
    {
        SheetCustomization s = new();
        SheetCustomization returned = s.Protect();

        Assert.IsTrue(ReferenceEquals(s, returned));
    }

    [TestMethod]
    public void Protect_Null_ShouldThrows_ArgumentNullException()
    {
        SheetCustomization s = null;
        Assert.ThrowsException<ArgumentNullException>(() => s.Protect());
    }
}
