using AwesomeExcel.BridgeNpoi;
using AwesomeExcel.Common.Models;

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
}
