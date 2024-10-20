using AwesomeExcel.Models;

namespace AwesomeExcel;

public class SheetCustomization
{
    public string Name { get; set; }

    public bool IsReadOnly { get; set; }

    public bool HasHeader { get; set; }

    public SheetStyle Style { get; set; }

    public Style HeaderStyle { get; set; }
}

public class SheetCustomization<TRows> : SheetCustomization
{

}