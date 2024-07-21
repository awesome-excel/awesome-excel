namespace AwesomeExcel;

public class CellFontStyleCustomization { }

public class CellFontStyleCustomization<T> : CellFontStyleCustomization
{
    public Func<T, string> Name { get; set; }
    public Func<T, Color?> Color { get; set; }
    public Func<T, short?> HeightInPoints { get; set; }
    public Func<T, bool?> IsBold { get; set; }
}
