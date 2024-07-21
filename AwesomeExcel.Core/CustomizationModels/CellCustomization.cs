namespace AwesomeExcel;

public interface ICellCustomization
{
    public CellStyleCustomization Style { get; }
}

public class CellCustomization<TProperty> : ICellCustomization
{
    public CellStyleCustomization<TProperty> Style { get; set; }

    CellStyleCustomization ICellCustomization.Style
    {
        get => Style; 
    }
}