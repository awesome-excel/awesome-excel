# AwesomeExcel

## Generate customizable Excel files via object mapping

```csharp
public void GenerateExcel()
{
    ExcelGenerator excel = new();
    
    // Get the invoices (or any data you need)
    List<Invoice> invoices = GetInvoices();
    
    // Generate the Excel file with some customizations
    MemoryStream file = excel.Generate(invoices, sheet =>
    {
        // Customize the entire sheet
        sheet
            .SetName("Client's invoices")
            .SetFontName("Aptos")
            .SetBordersColor(Color.Black);
        
        // Customize the header row
        sheet
            .HasHeader()
            .SetHeaderFontBold()
            .SetHeaderFontHeightInPoints(12)
            .SetHeaderHorizontalAlignment(HorizontalAlignment.Center)
            .SetHeaderFillForegroundColor(Color.Gray_25_Percent);
        
        // Customize only the specified column 
        sheet.Column(columns => columns.CreationDate)
            .SetName("Created on")
            .SetDateTimeFormat("dddd dd mmmm YYYY");
        
        // Customize the cells which amount is greater than 1000
        sheet.Cells(columns => columns.Amount)
            .SetFillForegroundColor(amount => amount > 1000 ? Color.LightGreen : null);
        });
    
    string fileName = "invoices.xlsx";
    WriteFile(file, fileName);
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
List<Person> actors = GetActors();
List<Invoice> invoices = GetInvoices();

MemoryStream file = excel.Generate(actors, invoices, (sheet1, sheet2) =>
{
    sheet1
        .SetName("Actors sheet")
        .SetFillForegroundColor(Color.LightBlue);
    
    // Customize the first sheet header row
    sheet1.HasHeader()
        .SetHeaderFillForegroundColor(Color.Blue)
        .SetHeaderBorderBottomColor(Color.Red)
        .SetVerticalAlignment(VerticalAlignment.Center);
    
    // Customize the specified column of the first sheet
    sheet1.Column(p => p.Surname)
        .SetName("Actor's surname")
        .SetHorizontalAlignment(HorizontalAlignment.Right);
    
    // Customize the specified column of the second sheet
    sheet2.Column(p => p.CreationDate)
        .SetDateTimeFormat("dd/mm/yyyy");
    
    // Customize the cells which amount is greater than 1000 on the second sheet
    sheet2.Cells(p => p.Amount)
        .SetFillForegroundColor(amount => amount >= 1500 ? Color.Green : Color.Red);
});
```

#### This library serves as:
- An easy-to-use abstraction layer over the NPOI engine (or any engine, as it allows developers to implement bridges for any engine).
- A quick solution for mapping objects to Excel files, which is useful for exporting data from your application. In this scenario, you would use Dapper or EntityFramework to retrieve the data from your database, resulting in a List of objects. You can then call AwesomeExcel's Generate() method directly with your dataset. Many companies already have "export to Excel" features in their reporting systems. However, these often rely on outdated, hard-to-read code that lacks customization options for colors, fonts, etc.

#### What about Aspose.Cell:
Aspose.Cell is a super-feature-rich library, and I’m ok with not matching its full functionality (which would take years to develop and might not be practical).
If you need those advanced features, Aspose.Cell is the best option. However, for small to medium-sized organizations looking to generate downloadable data reports with minimal code and easy customization, you can use my library or consider others available online, like SpreadSheetLite or ExcelMapper.


We currently don't support charts/formulas and reading Excel files.
