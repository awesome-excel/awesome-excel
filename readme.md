# AwesomeExcel

## Generate customizable Excel files via object mapping

```csharp
public void GenerateExcel()
{
    AwesomeExcel awesomeExcel = new();

    // Get the invoices (or any data you need)
    List<Invoice> invoices = GetInvoices();

    // Generate the Excel file with some customizations
    MemoryStream file = awesomeExcel.Generate(invoices, customization =>
    {
        // Customize the entire sheet
        customization.Sheet
            .SetName("Client's invoices")
            .SetFontName("Aptos")
            .SetBordersColor(Color.Black);

        // Customize the header row
        customization.Sheet
            .HasHeader()
            .SetHeaderFontBold()
            .SetHeaderFontHeightInPoints(12)
            .SetHeaderHorizontalAlignment(HorizontalAlignment.Center)
            .SetHeaderFillForegroundColor(Color.Gray_25_Percent);

        // Customize only the specified column 
        customization.Column(columns => columns.CreationDate)
            .SetName("Created on")
            .SetDateTimeFormat("dddd dd mmmm YYYY");

        // Customize only the cells which amount is greater than 1000
        customization.Cells(columns => columns.Amount)
            .SetFillForegroundColor(amount => amount > 1000 ? Color.LightGreen : null);
    });
}
```

### The result

![](https://i.imgur.com/hKpyML1.png)

### Sample data:

```csharp
public class Invoice
{
    public DateTime CreationDate { get; set; }
    public double Amount { get; set; }
}
```

```csharp
private List<Invoice> GetInvoices()
{
    return new List<Invoice>
    {
        { new() { Amount = 35, CreationDate = new DateTime(2016, 01, 01) } },
        { new() { Amount = 52, CreationDate = new DateTime(2016, 02, 01) } },
        { new() { Amount = 12312.3, CreationDate = new DateTime(2016, 03, 01) } },
        { new() { Amount = 3434.5654, CreationDate = new DateTime(2016, 04, 01) } },
        { new() { Amount = 234, CreationDate = new DateTime(2016, 05, 01) } },
        { new() { Amount = 12.3, CreationDate = new DateTime(2016, 06, 01) } },
        { new() { Amount = 35, CreationDate = new DateTime(2016, 06, 05) } },
        { new() { Amount = 7, CreationDate = new DateTime(2016, 06, 12) } },
        { new() { Amount = 3.567, CreationDate = new DateTime(2018, 01, 01) } },
        { new() { Amount = 8776, CreationDate = new DateTime(2019, 01, 01) } },
        { new() { Amount = 56.7, CreationDate = new DateTime(2020, 01, 01) } },
        { new() { Amount = 56.70, CreationDate = new DateTime(2020, 01, 02) } },
    };
}
```

### You can also create multiple sheets on a single workbook:

```csharp
AwesomeExcel awesomeExcel = new();

List<Person> actors = GetActors();
List<Invoice> invoices = GetInvoices();

MemoryStream file = awesomeExcel.Generate(actors, invoices, (SheetsCustomizer<Person, Invoice> customization) =>
{ 
    customization.Sheet1
        .SetName("Actors sheet")
        .HasHeader();

    customization.Column(customization.Sheet1, p => p.Name)
        .SetName("Actor's name")
        .SetFillForegroundColor(Color.Aqua);

    customization.Sheet2
        .SetName("Client's invoices")
        .SetFontName("Aptos")
        .SetBordersColor(Color.Black);

    customization.Column(customization.Sheet2, p => p.CreationDate)
        .SetName("Created on")
        .SetDateTimeFormat("dddd dd mmmm YYYY");
});
```

####
We currently don't support charts/formulas and reading Excel files.