using AwesomeExcel.Common.Models;
using AwesomeExcel.Generator;

namespace Tests.Generator;

[TestClass]
public class SheetFactoryTest
{
    private readonly List<Person> data = new()
    {
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) }
    };

    [TestMethod]
    public void Create_NullRows_ShouldThrow_ArgumentNullException()
    {
        SheetFactory factory = new();
        List<Person> rows = null;
        Assert.ThrowsException<ArgumentNullException>(() => factory.Create(rows, null, null, null));
    }

    [TestMethod]
    public void Create_RowsContainsNullElements_ShouldThrow_ArgumentNullException()
    {
        SheetFactory factory = new();
        List<Person> rows = new() { null };
        Assert.ThrowsException<InvalidOperationException>(() => factory.Create(rows, null, null, null));
    }

    [TestMethod]
    public void Create_Sheet_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        Assert.IsNotNull(sheet);
    }

    [TestMethod]
    public void Create_Columns_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);
        Assert.IsNotNull(sheet.Columns);
    }

    [TestMethod]
    public void Create_ColumnsCount_ShouldBe_Four()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        Assert.AreEqual(sheet.Columns.Count, 4);
    }

    [TestMethod]
    public void Create_Columns_ShouldHave_SpecifiedName()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        // First column
        Assert.AreEqual(sheet.Columns[0].Name, nameof(Person.Name));

        // Second column
        Assert.AreEqual(sheet.Columns[1].Name, nameof(Person.Surname));

        // Third column
        Assert.AreEqual(sheet.Columns[2].Name, nameof(Person.BirthDate));

        // Fourth column
        Assert.AreEqual(sheet.Columns[3].Name, nameof(Person.Age));
    }

    [TestMethod]
    public void Create_Columns_ShouldHave_RightType()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        // First column
        Assert.AreEqual(sheet.Columns[0].ColumnType, ColumnType.String);

        // Second column
        Assert.AreEqual(sheet.Columns[1].ColumnType, ColumnType.String);

        // Third column
        Assert.AreEqual(sheet.Columns[2].ColumnType, ColumnType.DateTime);

        // Fourth column
        Assert.AreEqual(sheet.Columns[3].ColumnType, ColumnType.Numeric);
    }

    [TestMethod]
    public void Create_Rows_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        // Rows
        Assert.IsNotNull(sheet.Rows);
        Assert.AreEqual(sheet.Rows.Count, 1);
    }

    [TestMethod]
    public void Create_FirstRow_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        Row firstRow = sheet.Rows[0];
        Assert.IsNotNull(firstRow);
    }

    [TestMethod]
    public void Create_Cells_ShouldNotBe_Null()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        // Cells
        Row firstRow = sheet.Rows[0];
        Assert.IsNotNull(firstRow.Cells);
        Assert.AreEqual(firstRow.Cells.Count, 4);
    }

    [TestMethod]
    public void Create_Cells_ShouldHave_SpecifiedValues()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);

        // Cells
        Row firstRow = sheet.Rows[0];
        IList<Cell> cells = firstRow.Cells;

        // First cell
        Assert.AreEqual(actual: cells[0].Value, expected: "Leonardo");

        // Second cell
        Assert.AreEqual(actual: cells[1].Value, expected: "DiCaprio");

        // Third cell
        Assert.AreEqual(actual: cells[2].Value, expected: new DateTime(1974, 11, 11));

        // Fourth cell
        int expectedAge = (int)(DateTime.Now - new DateTime(1974, 11, 11)).TotalDays / 365;
        Assert.AreEqual(cells[3].Value, expectedAge);
    }

    private class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age => (DateTime.Now.Date - BirthDate.Date).Days / 365;
    }
}
