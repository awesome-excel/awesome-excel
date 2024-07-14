using AwesomeExcel.Common.Models;
using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.Customization;

public static class WorkbookCustomizationExtension
{
    public static WorkbookCustomization SetFileType(this WorkbookCustomization workbook, FileType fileType)
    {
        workbook.FileType = fileType;
        return workbook;
    }
}
