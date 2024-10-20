using AwesomeExcel.Models;

namespace AwesomeExcel;

public static class WorkbookCustomizationExtension
{
    public static WorkbookCustomization SetFileType(this WorkbookCustomization workbook, FileType fileType)
    {
        workbook.FileType = fileType;
        return workbook;
    }
}
