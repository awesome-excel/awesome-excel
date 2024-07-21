namespace AwesomeExcel;

public class ColumnCustomization
{
    public string Name { get; set; }
    public bool Excluded { get; set; }
    public Style Style { get; set; }
}

public class ColumnCustomization<TValue> : ColumnCustomization
{

}

