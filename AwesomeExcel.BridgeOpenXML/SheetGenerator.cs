using OpenXml = DocumentFormat.OpenXml;

namespace AwesomeExcel.BridgeOpenXML;

public class SheetGenerator
{
    private readonly OpenXml.Packaging.WorkbookPart workbookPart;

    public SheetGenerator(OpenXml.Packaging.WorkbookPart workbookPart)
    {
        this.workbookPart = workbookPart;
    }

    public void GenerateSheet(AwesomeExcel.Models.Sheet item)
    {
        OpenXml.Packaging.WorksheetPart worksheetPart = workbookPart.AddNewPart<OpenXml.Packaging.WorksheetPart>();
        OpenXml.Spreadsheet.SheetData sheetData = CreateSheetData(item);
        worksheetPart.Worksheet = CreateWorkSheet(sheetData);

        uint sheetId = (uint)workbookPart.WorksheetParts.Count();
        string sheetName = item.Name ?? $"Sheet {sheetId}";

        if (IsSheetNameTaken(sheetName))
        {
            throw new InvalidOperationException();
        }

        OpenXml.Spreadsheet.Sheet sheet = new()
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName
        };

        workbookPart.Workbook.Sheets!.Append(sheet);
    }

    private bool IsSheetNameTaken(string sheetName)
    {
        return workbookPart.Workbook.Sheets.Any(sheet => ((OpenXml.Spreadsheet.Sheet)sheet).Name == sheetName);
    }

    private OpenXml.Spreadsheet.Worksheet CreateWorkSheet(OpenXml.Spreadsheet.SheetData sheetData)
    {
        var worksheet = new OpenXml.Spreadsheet.Worksheet();
        OpenXml.Spreadsheet.Columns columns1 = AutoSizeColumns.AutoSize(sheetData);

        worksheet.Append(columns1);
        worksheet.Append(sheetData);

        return worksheet;
    }

    private OpenXml.Spreadsheet.SheetData CreateSheetData(AwesomeExcel.Models.Sheet item)
    {
        OpenXml.Spreadsheet.SheetData data = new();

        foreach (AwesomeExcel.Models.Row row in item.Rows)
        {
            OpenXml.Spreadsheet.Row openXmlRow = CreateRow(row);
            data.Append(openXmlRow);
        }

        return data;
    }

    private OpenXml.Spreadsheet.Row CreateRow(AwesomeExcel.Models.Row row)
    {
        OpenXml.Spreadsheet.Row openXmlRow = new();

        foreach (AwesomeExcel.Models.Cell cell in row.Cells)
        {
            OpenXml.Spreadsheet.Cell openXmlCell = CreateCell(cell);
            openXmlRow.Append(openXmlCell);
        }

        return openXmlRow;
    }

    private OpenXml.Spreadsheet.Cell CreateCell(AwesomeExcel.Models.Cell cell)
    {
        ArgumentNullException.ThrowIfNull(cell);

        object cellValue = cell.Value;

        if (cellValue is null)
        {
            return new OpenXml.Spreadsheet.Cell()
            {
                DataType = OpenXml.Spreadsheet.CellValues.String,
                CellValue = new OpenXml.Spreadsheet.CellValue("")
            };
        }

        Type type = Nullable.GetUnderlyingType(cellValue.GetType()) ?? cellValue.GetType();

        if (IsNumber(type))
        {
            return new()
            {
                DataType = OpenXml.Spreadsheet.CellValues.Number,
                CellValue = new OpenXml.Spreadsheet.CellValue(cellValue.ToString())
            };
        }
        else if (type == typeof(DateTime) || type == typeof(DateTime?))
        {
            // https://stackoverflow.com/questions/2792304/how-to-insert-a-date-to-an-open-xml-worksheet

            if (type == typeof(DateTime?))
            {
                cellValue = ((DateTime?)cellValue).Value;
            }

            DateTime valueDate = (DateTime)cellValue;
            string valueString = valueDate.ToOADate().ToString();

            return new()
            {
                DataType = OpenXml.Spreadsheet.CellValues.Number,
                CellValue = new OpenXml.Spreadsheet.CellValue(valueString)
            };
        }
        else
        {
            return new()
            {
                DataType = OpenXml.Spreadsheet.CellValues.String,
                CellValue = new OpenXml.Spreadsheet.CellValue(cellValue.ToString())
            };
        }

        static bool IsNumber(Type t)
        {
            return t == typeof(short)
                || t == typeof(ushort)
                || t == typeof(int)
                || t == typeof(uint)
                || t == typeof(long)
                || t == typeof(ulong)
                || t == typeof(float)
                || t == typeof(double)
                || t == typeof(decimal);
        }
    }

    private OpenXml.Spreadsheet.Stylesheet CreateStyleSheet()
    {
        OpenXml.Spreadsheet.Stylesheet stylesheet = new();
        return stylesheet;
    }

    public void GenerateStyle()
    {
        OpenXml.Packaging.WorkbookStylesPart stylesPart = workbookPart.AddNewPart<OpenXml.Packaging.WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStyleSheet();
        stylesPart.Stylesheet.Save();
    }
}