using OpenXml = DocumentFormat.OpenXml;

namespace AwesomeExcel.BridgeOpenXML;

public class OpenXmlFileGenerator : Core.Services.FileGenerator<AwesomeExcel.Models.Workbook>
{
    protected override Models.Workbook Convert(Models.Workbook workbook)
    {
        return workbook;
    }

    protected override MemoryStream Write(AwesomeExcel.Models.Workbook workbook)
    {
        MemoryStream memoryStream = new();

        // https://jason-ge.medium.com/create-excel-using-openxml-in-net-6-3b601ddf48f7

        using OpenXml.Packaging.SpreadsheetDocument document = OpenXmlHelper.CreateDocument(memoryStream);

        SheetGenerator sheetGenerator = new(document.WorkbookPart);

        foreach (AwesomeExcel.Models.Sheet item in workbook.Sheets)
        {
            sheetGenerator.GenerateSheet(item);
        }

        sheetGenerator.GenerateStyle();

        return memoryStream;
    }
}
