using AwesomeExcel;
using AwesomeExcel.Core.Services;

namespace Tests.Generator;

[TestClass]
public class WorkbookFactoryTest
{
    private readonly List<Person> data = new()
    {
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) }
    };

    [TestMethod]
    public void Create_NullSheets_ShouldThrow_ArgumentNullException()
    {
        WorkbookFactory factory = new();
        List<Sheet> sheets = null;
        Assert.ThrowsException<ArgumentNullException>(() => factory.Create(sheets, customization: null));
    }

    [TestMethod]
    public void Create_SheetsListContainsNullElements_ShouldThrow_InvalidOperationException()
    {
        WorkbookFactory factory = new();
        List<Sheet> sheets = new() { null };
        Assert.ThrowsException<InvalidOperationException>(() => factory.Create(sheets, customization: null));
    }

    [TestMethod]
    public void Create_NoSheets_ShouldThrow_InvalidOperationException()
    {
        WorkbookFactory factory = new();
        List<Sheet> sheets = new() { };
        Assert.ThrowsException<InvalidOperationException>(() => factory.Create(sheets, customization: null));
    }

    [TestMethod]
    public void Create_Workbook_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        WorkbookFactory workbookFactory = new();

        Sheet sheet = factory.Create(data, null, null, null);
        Workbook wb = workbookFactory.Create(new[] { sheet }, null);

        Assert.IsNotNull(wb);
    }

    [TestMethod]
    public void Create_Sheets_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        WorkbookFactory workbookFactory = new();

        Sheet sheet = factory.Create(data, null, null, null);
        Workbook wb = workbookFactory.Create(new[] { sheet }, null);

        Assert.IsNotNull(wb.Sheets);
    }

    [TestMethod]
    public void Create_SheetsCount_ShouldBe_One()
    {
        SheetFactory factory = new();
        WorkbookFactory workbookFactory = new();

        Sheet sheet = factory.Create(data, null, null, null);
        Workbook wb = workbookFactory.Create(new[] { sheet }, null);

        Assert.AreEqual(wb.Sheets.Count, 1);
    }

    private class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age => (DateTime.Now.Date - BirthDate.Date).Days / 365;
    }
}
