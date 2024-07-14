namespace AwesomeExcel.Customization.Models;

public class CellCustomization
{

}

public class CellCustomization<TProperty> : CellCustomization
{
    public CellStyleCustomization<TProperty> Style { get; set; }
}