using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization;
using AwesomeExcel.Customization.Models;
using AwesomeExcel.Generator;
using System.Reflection;

namespace Tests.Generator;

[TestClass]
public class SheetFactory_ColumnsCustomizationTest
{
    private readonly List<Person> data = new()
    {
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
        new Person() { Name = "Leonardo", Surname = "DiCaprio", BirthDate = new DateTime(1974, 11, 11) },
    };

    [TestMethod]
    public void Create_CustomColumnsName_ShouldReturn_GivenCustomNames()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, GetCustomizedColumn(), null);

        Assert.AreEqual(sheet.Columns.Count, 4);
        Assert.AreEqual(sheet.Columns[0].Name, "Actor's name");
        Assert.AreEqual(sheet.Columns[1].Name, "Actor's surname");
        Assert.AreEqual(sheet.Columns[2].Name, "Actor's date of birth");
        Assert.AreEqual(sheet.Columns[3].Name, "Actor's age");
    }

    [TestMethod]
    public void CreateSheet_NullChecks()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, GetCustomizedColumn(), null);

        Assert.IsNotNull(sheet);
        Assert.IsNotNull(sheet.Columns);
        Assert.IsNotNull(sheet.Rows);
        Assert.IsNull(sheet.Style);
        Assert.IsNull(sheet.HeaderStyle);

        Column column0 = sheet.Columns[0];
        Assert.IsNotNull(column0.Style);

        Column column1 = sheet.Columns[1];
        Assert.IsNotNull(column1.Style);

        Column column2 = sheet.Columns[2];
        Assert.IsNotNull(column2.Style?.FontStyle);

        Column column3 = sheet.Columns[3];
        Assert.IsNotNull(column3.Style?.FontStyle);
    }

    [TestMethod]
    public void CreateSheet_Column0_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, GetCustomizedColumn(), null);

        Column column0 = sheet.Columns[0];
        Assert.AreEqual(column0.Style.HorizontalAlignment, HorizontalAlignment.Right);
    }

    [TestMethod]
    public void CreateSheet_Column1_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, GetCustomizedColumn(), null);

        Column column1 = sheet.Columns[1];
        Assert.AreEqual(column1.Style.FillForegroundColor, Color.Blue);
    }

    [TestMethod]
    public void CreateSheet_Column2_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, GetCustomizedColumn(), null);

        Column column2 = sheet.Columns[2];
        Assert.IsTrue(column2.Style.FontStyle.IsBold);
        Assert.AreEqual(column2.Style.FontStyle.Color, Color.Red);
    }

    [TestMethod]
    public void CreateSheet_Column3_CustomStyle_ShouldReturn_CustomizedStyle()
    {
        SheetFactory factory = new();
        Sheet sheet = factory.Create(data, null, GetCustomizedColumn(), null);

        Column column3 = sheet.Columns[3];
        Assert.AreEqual(actual: column3.Style.BorderBottomColor, expected: Color.Green);
        Assert.AreEqual(actual: column3.Style.BorderLeftColor, expected: Color.IceBlue);
        Assert.AreEqual(actual: column3.Style.BorderRightColor, expected: Color.Lime);
        Assert.AreEqual(actual: column3.Style.BorderTopColor, expected: Color.Red);
        Assert.AreEqual(actual: column3.Style.FillForegroundColor, expected: Color.Ivory);
        //Assert.AreEqual(actual: column3.Style.FillPattern, expected: FillPattern.SolidForeground);
        Assert.AreEqual(actual: column3.Style.HorizontalAlignment, expected: HorizontalAlignment.Right);
        Assert.AreEqual(actual: column3.Style.VerticalAlignment, expected: VerticalAlignment.Bottom);
        //Assert.AreEqual(actual: fourhColumn.Style.DateTimeFormat, expected: "a5b2");

        Assert.AreEqual(actual: column3.Style.FontStyle.Color, expected: Color.Yellow);
        Assert.AreEqual(actual: column3.Style.FontStyle.HeightInPoints, expected: (short)12);
        Assert.AreEqual(actual: column3.Style.FontStyle.IsBold, expected: null);
        Assert.AreEqual(actual: column3.Style.FontStyle.Name, expected: "AwesomeExcel");
    }

    private class Person
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age => (DateTime.Now.Date - BirthDate.Date).Days / 365;
    }

    public Dictionary<PropertyInfo, ColumnCustomization> GetCustomizedColumn()
    {
        return new Dictionary<PropertyInfo, ColumnCustomization>
        {
            { typeof(Person).GetProperty(nameof(Person.Name)), GetColumnName() },
            { typeof(Person).GetProperty(nameof(Person.Surname)), GetColumnSurname() },
            { typeof(Person).GetProperty(nameof(Person.BirthDate)), GetColumnBirthDate() },
            { typeof(Person).GetProperty(nameof(Person.Age)), GetColumnAge() },
        };
    }

    public ColumnCustomization GetColumn(PropertyInfo pi)
    {
        return pi.Name switch
        {
            nameof(Person.Name) => GetColumnName(),
            nameof(Person.Surname) => GetColumnSurname(),
            nameof(Person.BirthDate) => GetColumnBirthDate(),
            nameof(Person.Age) => GetColumnAge(),
            _ => throw new KeyNotFoundException()
        };
    }

    private ColumnCustomization GetColumnAge()
    {
        ColumnCustomization customizedColumn = new ColumnCustomization();

        customizedColumn
            .SetName("Actor's age")
            .SetBorderLeftColor(Color.IceBlue)
            .SetBorderRightColor(Color.Lime)
            .SetBorderTopColor(Color.Red)
            .SetBorderBottomColor(Color.Green)
            .SetFontHeightInPoints(12)
            .SetFontColor(Color.Yellow)
            .SetHorizontalAlignment(HorizontalAlignment.Right)
            .SetVerticalAlignment(VerticalAlignment.Bottom)
            .SetFillForegroundColor(Color.Ivory)
            .SetFontName("AwesomeExcel");


        return customizedColumn;
    }

    private ColumnCustomization GetColumnBirthDate()
    {
        ColumnCustomization customizedColumn = new();

        customizedColumn.SetName("Actor's date of birth");
        customizedColumn.SetStyle(s =>
        {
            s.FontStyle.IsBold = true;
            s.FontStyle.Color = Color.Red;
        });

        return customizedColumn;
    }

    private ColumnCustomization GetColumnSurname()
    {
        ColumnCustomization customizedColumn = new();

        customizedColumn.Name = "Actor's surname";
        customizedColumn.Style = new()
        {
            FillForegroundColor = Color.Blue
        };


        return customizedColumn;
    }

    private ColumnCustomization GetColumnName()
    {
        ColumnCustomization _customizedColumn = new();

        _customizedColumn
            .SetName("Actor's name")
            .SetStyle(s => s.HorizontalAlignment = HorizontalAlignment.Right);

        return _customizedColumn;
    }
}
