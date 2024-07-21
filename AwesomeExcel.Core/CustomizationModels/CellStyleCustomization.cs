namespace AwesomeExcel;

public class CellStyleCustomization { }

public class CellStyleCustomization<T> : CellStyleCustomization
{
    public Func<T, Color?> BorderTopColor { get; set; }
    public Func<T, Color?> BorderBottomColor { get; set; }
    public Func<T, Color?> BorderLeftColor { get; set; }
    public Func<T, Color?> BorderRightColor { get; set; }
    public Func<T, Color?> FillForegroundColor { get; set; }
    public Func<T, FillPattern?> FillPattern { get; set; }
    public Func<T, string> DateTimeFormat { get; set; }
    public Func<T, HorizontalAlignment?> HorizontalAlignment { get; set; }
    public Func<T, VerticalAlignment?> VerticalAlignment { get; set; }

    public CellFontStyleCustomization<T> FontStyle { get; set; }
}
