using AwesomeExcel.BridgeNpoi;
using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization;
using AwesomeExcel.Customization.Services;
using AwesomeExcel.Generator;
using Tests.IntegrationTests;

namespace Tests.BridgeNpoi;

[TestClass]
public class WorkbookConverterTest
{
    [TestMethod]
    public void Convert_Null_ThrowsException()
    {
        WorkbookConverter converter = new();
        Assert.ThrowsException<ArgumentNullException>(() => converter.Convert(null));
    }

    [TestMethod]
    public void Convert_EmptyWorkbook_ThrowsException()
    {
        Workbook excelWorkbook = new();
        WorkbookConverter converter = new();

        Assert.ThrowsException<InvalidOperationException>(() => converter.Convert(excelWorkbook));
    }

    [TestMethod]
    public void Convert_NoSheets_ThrowsException()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[0]
        };
        WorkbookConverter converter = new();

        Assert.ThrowsException<InvalidOperationException>(() => converter.Convert(excelWorkbook));
    }

    [TestMethod]
    public void Convert_EmptySheet_ThrowsException()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1] { null }
        };
        WorkbookConverter converter = new();

        Assert.ThrowsException<InvalidOperationException>(() => converter.Convert(excelWorkbook));
    }

    [TestMethod]
    public void TestMethod5()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1] { new() }
        };

        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);

        Assert.IsNotNull(npoiWorkbook);
        Assert.AreEqual(actual: npoiWorkbook.NumberOfSheets, expected: 1);
    }

    [TestMethod]
    public void TestMethod50()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1] { new() }
        };

        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.IsNotNull(npoiSheet);
    }

    [TestMethod]
    public void TestMethod51()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1] { new() }
        };

        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.AreEqual(actual: npoiSheet.PhysicalNumberOfRows, expected: 0);
    }

    [TestMethod]
    public void TestMethod6()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1]
            {
                new()
                {
                    Name = "Daniel LaRusso",
                    Rows = new Row[3]
                    {
                        null,
                        null,
                        null
                    }
                }
            }
        };
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.AreEqual(expected: "Daniel LaRusso", actual: npoiSheet.SheetName);
        Assert.AreEqual(expected: 3, actual: npoiSheet.PhysicalNumberOfRows);
    }

    [TestMethod]
    public void TestMethod7()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1]
            {
                new()
                {
                    Name = "Daniel LaRusso",
                    Rows = new Row[1]
                    {
                        new()
                        {
                            Style = null,
                            Cells = null
                        }
                    }
                }
            }
        };
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.AreEqual(expected: "Daniel LaRusso", actual: npoiSheet.SheetName);
        Assert.AreEqual(expected: 1, actual: npoiSheet.PhysicalNumberOfRows);
    }

    [TestMethod]
    public void TestMethod8()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1]
            {
                new()
                {
                    Name = "Daniel LaRusso",
                    Rows = new Row[1]
                    {
                        new()
                        {
                            Style = new()
                            {
                                FontStyle = null
                            },
                            Cells = new Cell[3]
                            {
                                null,
                                null,
                                null
                            }
                        }
                    }
                }
            }
        };
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.IsNotNull(npoiWorkbook.GetSheetAt(0));
        Assert.AreEqual(expected: "Daniel LaRusso", actual: npoiSheet.SheetName);
        Assert.AreEqual(expected: 1, actual: npoiSheet.PhysicalNumberOfRows);
    }

    [TestMethod]
    public void TestMethod9()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1]
            {
                new()
                {
                    Name = "Daniel LaRusso",
                    Rows = new Row[1]
                    {
                        new()
                        {
                            Style = new()
                            {
                                FontStyle = null
                            },
                            Cells = new Cell[1]
                            {
                                new()
                            }
                        }
                    }
                }
            }
        };
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.IsNotNull(npoiSheet);
        Assert.AreEqual(expected: "Daniel LaRusso", actual: npoiSheet.SheetName);
        Assert.AreEqual(expected: 1, actual: npoiSheet.PhysicalNumberOfRows);
        Assert.AreEqual(expected: 1, actual: npoiSheet.GetRow(0).PhysicalNumberOfCells);
    }

    [TestMethod]
    public void TestMethod10()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new Sheet[1]
            {
                new()
                {
                    Name = "Daniel LaRusso",
                    Rows = new Row[1]
                    {
                        new()
                        {
                            Style = new()
                            {
                                FontStyle = null
                            },
                            Cells = new Cell[2]
                            {
                                new()
                                {
                                    Value = null,
                                    Style = null
                                },
                                new()
                                {
                                    Value = "Mr. Miyagi",
                                    Style = null
                                }
                            }
                        }
                    }
                }
            }
        };
        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.AreEqual(expected: 2, actual: npoiSheet.GetRow(0).PhysicalNumberOfCells);
        Assert.AreEqual(expected: "", actual: npoiSheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.AreEqual(expected: "Mr. Miyagi", actual: npoiSheet.GetRow(0).GetCell(1).StringCellValue);
    }

    [TestMethod]
    public void TestMethod11()
    {
        Workbook excelWorkbook = new()
        {
            Sheets = new List<Sheet>()
        };

        for (int z = 0; z < 7; z++)
        {
            Sheet newSheet = new()
            {
                Name = z.ToString(),
                Rows = new List<Row>()
            };

            for (int i = 0; i < 5000; i++)
            {
                Row newRow = new()
                {
                    Cells = new List<Cell>()
                };

                for (int y = 0; y < 20; y++)
                {
                    Cell newCell = new()
                    {
                        Value = y
                    };
                    newRow.Cells.Add(newCell);
                }

                newSheet.Rows.Add(newRow);
            }

            excelWorkbook.Sheets.Add(newSheet);
        }

        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
    }

    [TestMethod]
    public void TestMethod12()
    {
        Style redBackground = new()
        {
            FillForegroundColor = Color.Red
        };
        Style greenBackground = new()
        {
            FillForegroundColor = Color.Green
        };
        Style blueBackground = new()
        {
            FillForegroundColor = Color.Blue
        };
        Style boldFont = new()
        {
            FontStyle = new()
            {
                IsBold = true
            }
        };

        var rows = new List<Row>();

        for (int i = 0; i < 10; i++)
        {
            List<Cell> cells = Enumerable.Range(0, 10)
                .Select(i => new Cell() { Style = i == 0 ? boldFont : null })
                .ToList();

            Style s = (i % 3) switch
            {
                0 => redBackground,
                1 => greenBackground,
                2 => blueBackground,
                _ => null
            };

            Row row = new()
            {
                Cells = cells,
                Style = s
            };
            rows.Add(row);
        }

        Sheet sheet = new()
        {
            Rows = rows
        };

        Workbook excelWorkbook = new()
        {
            Sheets = new List<Sheet>
            {
                sheet
            }
        };

        WorkbookConverter converter = new();
        NPOI.SS.UserModel.IWorkbook npoiWorkbook = converter.Convert(excelWorkbook);
        NPOI.SS.UserModel.ISheet npoiSheet = npoiWorkbook.GetSheetAt(0);

        Assert.AreEqual(expected: rows.Count, actual: npoiSheet.PhysicalNumberOfRows);

        NPOI.SS.UserModel.IRow row0 = npoiSheet.GetRow(0);
        NPOI.SS.UserModel.IRow row1 = npoiSheet.GetRow(1);
        NPOI.SS.UserModel.IRow row2 = npoiSheet.GetRow(2);
        NPOI.SS.UserModel.IRow row3 = npoiSheet.GetRow(3);
        NPOI.SS.UserModel.IRow row4 = npoiSheet.GetRow(4);
        NPOI.SS.UserModel.IRow row5 = npoiSheet.GetRow(5);
        NPOI.SS.UserModel.IRow row6 = npoiSheet.GetRow(6);
        NPOI.SS.UserModel.IRow row7 = npoiSheet.GetRow(7);
        NPOI.SS.UserModel.IRow row8 = npoiSheet.GetRow(8);
        NPOI.SS.UserModel.IRow row9 = npoiSheet.GetRow(9);

        Assert.IsNotNull(row0.RowStyle);
        Assert.AreEqual(expected: (short)Color.Red, actual: row0.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row0.RowStyle.FillPattern);

        Assert.IsNotNull(row3.RowStyle);
        Assert.AreEqual(expected: (short)Color.Red, actual: row3.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row3.RowStyle.FillPattern);

        Assert.IsNotNull(row6.RowStyle);
        Assert.AreEqual(expected: (short)Color.Red, actual: row6.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row6.RowStyle.FillPattern);

        Assert.IsNotNull(row9.RowStyle);
        Assert.AreEqual(expected: (short)Color.Red, actual: row9.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row9.RowStyle.FillPattern);

        Assert.IsNotNull(row1.RowStyle);
        Assert.AreEqual(expected: (short)Color.Green, actual: row1.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row1.RowStyle.FillPattern);

        Assert.IsNotNull(row4.RowStyle);
        Assert.AreEqual(expected: (short)Color.Green, actual: row4.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row4.RowStyle.FillPattern);

        Assert.IsNotNull(row7.RowStyle);
        Assert.AreEqual(expected: (short)Color.Green, actual: row7.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row7.RowStyle.FillPattern);

        Assert.IsNotNull(row2.RowStyle);
        Assert.AreEqual(expected: (short)Color.Blue, actual: row2.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row2.RowStyle.FillPattern);

        Assert.IsNotNull(row5.RowStyle);
        Assert.AreEqual(expected: (short)Color.Blue, actual: row5.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row5.RowStyle.FillPattern);

        Assert.IsNotNull(row8.RowStyle);
        Assert.AreEqual(expected: (short)Color.Blue, actual: row8.RowStyle.FillForegroundColor);
        Assert.AreEqual(expected: NPOI.SS.UserModel.FillPattern.SolidForeground, actual: row8.RowStyle.FillPattern);

        Assert.AreEqual(expected: true, actual: row0.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row1.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row2.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row3.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row4.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row5.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row6.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row7.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row8.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
        Assert.AreEqual(expected: true, actual: row9.Cells[0].CellStyle.GetFont(npoiWorkbook).IsBold);
    }

    [TestMethod]
    public void Test12()
    {
        List<Person> rows = GetActors();
        SheetCustomizer<Person> customizer = new();
        customizer.HasHeader();
        customizer.Column(person => person.Name).SetName("person name");

        Sheet sheet = new SheetFactory().Create(rows, customizer, customizer?.GetColumns(), customizer?.GetCells());
        Workbook workbook = new WorkbookFactory().Create(new Sheet[1] { sheet }, customizer?.Workbook);

        NPOI.SS.UserModel.IWorkbook npoiWorkbook = new WorkbookConverter().Convert(workbook);
    }

    private List<Person> GetActors()
    {
        return new List<Person>
        {
            { new() { Name =  "Caroline", Surname = "Aaron", BirthDate = DateTime.Parse("1952-08-07") } },
            { new() { Name =  "Victor", Surname = "Aaron", BirthDate = DateTime.Parse("1956-09-11") } },
            { new() { Name =  "Diego", Surname = "Abatantuono", BirthDate = DateTime.Parse("1955-05-20") } },
            { new() { Name =  "Andrew", Surname = "Abeita", BirthDate = DateTime.Parse("1981-07-11") } },
            { new() { Name =  "Jon", Surname = "Abrahams", BirthDate = DateTime.Parse("1977-10-29") } },
            { new() { Name =  "Stefano", Surname = "Accorsi", BirthDate = DateTime.Parse("1971-03-02") } },
            { new() { Name =  "Dean", Surname = "Acheson", BirthDate = DateTime.Parse("1893-04-11") } },
            { new() { Name =  "Josh", Surname = "Ackerman", BirthDate = DateTime.Parse("1977-03-23") } },
            { new() { Name =  "Joss", Surname = "Ackland", BirthDate = DateTime.Parse("1928-02-29") } },
            { new() { Name =  "Jay", Surname = "Acovone", BirthDate = DateTime.Parse("1955-08-20") } },
            { new() { Name =  "Deb", Surname = "Adair", BirthDate = DateTime.Parse("1966-04-22") } },
            { new() { Name =  "Enid-Raye", Surname = "Adams", BirthDate = DateTime.Parse("1973-06-16") } },
            { new() { Name =  "Jacob", Surname = "Adams", BirthDate = DateTime.Parse("1975-07-04") } },
            { new() { Name =  "Mario", Surname = "Adorf", BirthDate = DateTime.Parse("1930-09-08") } },
            { new() { Name =  "Ben", Surname = "Affleck", BirthDate = DateTime.Parse("1972-08-15") } },
            { new() { Name =  "Casey", Surname = "Affleck", BirthDate = DateTime.Parse("1975-08-12") } },
            { new() { Name =  "Spiro", Surname = "Agnew", BirthDate = DateTime.Parse("1918-11-09") } },
            { new() { Name =  "Antonio", Surname = "Agri", BirthDate = DateTime.Parse("1932-05-05") } },
            { new() { Name =  "Jenny", Surname = "Agutter", BirthDate = DateTime.Parse("1952-12-20") } },
            { new() { Name =  "Betsy", Surname = "Aidem", BirthDate = DateTime.Parse("1957-10-28") } },
            { new() { Name =  "Liam", Surname = "Aiken", BirthDate = DateTime.Parse("1990-01-07") } },
            { new() { Name =  "Troy", Surname = "Aikman", BirthDate = DateTime.Parse("1966-11-21") } },
            { new() { Name =  "Kacey", Surname = "Ainsworth", BirthDate = DateTime.Parse("1970-10-19") } },
            { new() { Name =  "Holly", Surname = "Aird", BirthDate = DateTime.Parse("1969-05-18") } },
            { new() { Name =  "Lucy", Surname = "Akhurst", BirthDate = DateTime.Parse("1975-11-18") } },
            { new() { Name =  "Amy", Surname = "Alcott", BirthDate = DateTime.Parse("1956-02-22") } },
            { new() { Name =  "Alan", Surname = "Alda", BirthDate = DateTime.Parse("1936-01-28") } },
            { new() { Name =  "Tom", Surname = "Aldredge", BirthDate = DateTime.Parse("1928-02-28") } },
            { new() { Name =  "Buzz", Surname = "Aldrin", BirthDate = DateTime.Parse("1930-01-20") } },
            { new() { Name =  "Henry", Surname = "Alessandroni", BirthDate = DateTime.Parse("1959-05-26") } },
            { new() { Name =  "Art", Surname = "Alexakis", BirthDate = DateTime.Parse("1962-04-12") } },
            { new() { Name =  "Jane", Surname = "Alexander", BirthDate = DateTime.Parse("1939-10-28") } },
            { new() { Name =  "Jason", Surname = "Alexander", BirthDate = DateTime.Parse("1959-09-23") } },
            { new() { Name =  "Khandi", Surname = "Alexander", BirthDate = DateTime.Parse("1957-09-04") } },
            { new() { Name =  "Adam", Surname = "Alexi-Malle", BirthDate = DateTime.Parse("1964-09-24") } },
            { new() { Name =  "Hans", Surname = "Alfredson", BirthDate = DateTime.Parse("1931-06-28") } },
            { new() { Name =  "Mary", Surname = "Alice", BirthDate = DateTime.Parse("1941-12-03") } },
            { new() { Name =  "Debbie", Surname = "Allen", BirthDate = DateTime.Parse("1950-01-16") } }
        };
    }
}
