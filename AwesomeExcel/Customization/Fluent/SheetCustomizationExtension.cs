using AwesomeExcel.Customization.Models;

namespace AwesomeExcel.Customization;

public static class SheetCustomizationExtension
{
    public static SheetCustomization SetName(this SheetCustomization sheetCustomization, string name)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        sheetCustomization.Name = name;
        return sheetCustomization;
    }

    public static SheetCustomization Protect(this SheetCustomization sheetCustomization)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        sheetCustomization.IsReadOnly = true;
        return sheetCustomization;
    }

    public static SheetCustomization HasHeader(this SheetCustomization sheetCustomization, bool hasHeader = true)
    {
        if (sheetCustomization is null)
        {
            throw new ArgumentNullException(nameof(sheetCustomization));
        }

        sheetCustomization.HasHeader = hasHeader;
        return sheetCustomization;
    }
}
