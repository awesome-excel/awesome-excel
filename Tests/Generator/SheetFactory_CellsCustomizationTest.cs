using AwesomeExcel;
using AwesomeExcel.Models;
using AwesomeExcel.Core.Services;
using System.Reflection;

namespace Tests.Generator;

[TestClass]
public class SheetFactory_CellsCustomizationTest
{
    private readonly List<Person> data = new()
    {
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "da Vinci", BirthDate = new DateTime(1452, 04, 15) },
        new Person() { Name = "Leonardo", Surname = "Fibonacci", BirthDate = new DateTime(1170, 09, 01) },
    };

    [TestMethod]
    public void CreateSheet_NullCellsCustomization_Successfull()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, null);
    }

    [TestMethod]
    public void CreateSheet_EmptyCellsCustomization_Successfull()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, new Dictionary<PropertyInfo, ICellCustomization>());
    }

    [TestMethod]
    public void CreateSheet_ColumnName_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        foreach (Row row in sheet.Rows)
        {
            const int columnName = 0;
            Cell cell = row.Cells.ElementAt(columnName);

            Assert.AreEqual(cell.Style.HorizontalAlignment, HorizontalAlignment.Right);
        }
    }

    [TestMethod]
    public void CreateSheet_ColumnSurname_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        foreach (Row row in sheet.Rows)
        {
            const int columnSurname = 1;
            Cell cell = row.Cells.ElementAt(columnSurname);
            string surname = (string)cell.Value;

            Color? expected = surname == "da Vinci"
                ? Color.Blue
                : surname == "DiCaprio"
                    ? Color.Green
                    : null;

            Color? actual = cell.Style.FillForegroundColor;

            Assert.AreEqual(actual, expected);
        }
    }

    [TestMethod]
    public void CreateSheet_ColumnBirthDate_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        foreach (Row row in sheet.Rows)
        {
            const int columnBirthDate = 2;
            Cell cell = row.Cells.ElementAt(columnBirthDate);
            var birthDate = (DateTime)cell.Value;

            bool expectedBold = birthDate.Month <= 6;
            Assert.AreEqual(actual: cell.Style.FontStyle.IsBold, expected: expectedBold);

            Color expectedFontColor = birthDate < new DateTime(1950, 1, 1) ? Color.Red : Color.SkyBlue;
            Assert.AreEqual(actual: cell.Style.FontStyle.Color, expected: expectedFontColor);
        }
    }

    [TestMethod]
    public void CreateSheet_ColumnAge_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, null, GetCustomizedCells());

        foreach (Row row in sheet.Rows)
        {
            const int columnAge = 3;
            Cell cell = row.Cells.ElementAt(columnAge);
            var age = (int)cell.Value;

            VerticalAlignment? expectedVerticalAlignment = age > 500 ? VerticalAlignment.Bottom : null;
            Assert.AreEqual(actual: cell.Style.VerticalAlignment, expected: expectedVerticalAlignment);

            short expectedHeight = (short)(age % 32767);
            Assert.AreEqual(actual: cell.Style.FontStyle.HeightInPoints, expected: expectedHeight);
        }
    }

    private class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age => (DateTime.Now.Date - BirthDate.Date).Days / 365;
    }

    public Dictionary<PropertyInfo, ICellCustomization> GetCustomizedCells()
    {
        return new Dictionary<PropertyInfo, ICellCustomization>
        {
            { typeof(Person).GetProperty(nameof(Person.Name)), GetColumnName() },
            { typeof(Person).GetProperty(nameof(Person.Surname)), GetColumnSurname() },
            { typeof(Person).GetProperty(nameof(Person.BirthDate)), GetColumnBirthDate() },
            { typeof(Person).GetProperty(nameof(Person.Age)), GetColumnAge() },
        };
    }

    private ICellCustomization GetColumnAge()
    {
        return new CellCustomization<int>()
            .SetFontHeightInPoints(value => (short)(value % 32767))
            .SetVerticalAlignment(value => value > 500 ? VerticalAlignment.Bottom : null);
    }

    private ICellCustomization GetColumnBirthDate()
    {
        DateTime dt1950 = new(1950, 1, 1);

        return new CellCustomization<DateTime>()
            .SetFontBold(value => value.Month <= 6)
            .SetFontColor(value => value < dt1950 ? Color.Red : Color.SkyBlue);
    }

    private ICellCustomization GetColumnSurname()
    {
        CellCustomization<string> customizedColumn = new();
        customizedColumn.SetFillForegroundColor(value =>
        {
            if (value == "da Vinci")
                return Color.Blue;

            if (value == "DiCaprio")
                return Color.Green;

            return null;
        });
        return customizedColumn;
    }

    private CellCustomization<string> GetColumnName()
    {
        return new CellCustomization<string>()
            .SetHorizontalAlignment(s => HorizontalAlignment.Right);
    }
}
