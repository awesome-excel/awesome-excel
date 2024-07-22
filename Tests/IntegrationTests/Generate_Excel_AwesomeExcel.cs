using AwesomeExcel;

namespace Tests.IntegrationTests;

[TestClass]
public class Generate_Excel_AwesomeExcel
{
    [TestMethod]
    public void Generate_Excel_Actors()
    {
        ExcelGenerator awesomeExcel = new();
        List<Person> actors = GetActors();

        MemoryStream file = awesomeExcel.Generate(actors, (SheetCustomizer<Person> sheet) =>
        {
            sheet.Workbook.SetFileType(FileType.Xlsx);

            sheet.SetName("Sheet's name test")
                .SetFillForegroundColor(Color.LightBlue)
                .SetHeaderFillForegroundColor(Color.Blue)
                .SetHeaderBorderBottomColor(Color.Red)
                .SetVerticalAlignment(VerticalAlignment.Center);

            sheet.Column(p => p.Name)
                .SetName("Actor's name")
                .SetStyle(s => s.FillForegroundColor = Color.Aqua);

            sheet.Column(p => p.Surname)
                .SetName("Actor's surname")
                .SetHorizontalAlignment(HorizontalAlignment.Right);

            sheet.Column(p => p.BirthDate)
                .SetStyle(s => s.DateTimeFormat = "dd/mm/yyyy");

            sheet.Cells(p => p.BirthDate)
                .SetFillForegroundColor(birthDate => birthDate.HasValue && birthDate.Value.Month == 3 ? Color.Red : null);
        });

        string fileName = nameof(Generate_Excel_Actors) + ".xlsx";
        WriteFile(file, fileName);
    }

    [TestMethod]
    public void Generate_Excel_Actors_Invoices()
    {
        ExcelGenerator excel = new();
        List<Person> actors = GetActors();
        List<Invoice> invoices = GetInvoices();

        MemoryStream file = excel.Generate(actors, invoices, (SheetCustomizer<Person> sheet1, SheetCustomizer<Invoice> sheet2) =>
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

        string fileName = nameof(Generate_Excel_Actors_Invoices) + ".xlsx";
        WriteFile(file, fileName);
    }

    [TestMethod]
    public void Generate_Excel_Actors_Invoices_NoCustomization()
    {
        ExcelGenerator awesomeExcel = new();
        List<Person> people = GetActors();
        List<Invoice> invoices = GetInvoices();

        MemoryStream file = awesomeExcel.Generate(people, invoices);

        string fileName = nameof(Generate_Excel_Actors_Invoices_NoCustomization) + ".xlsx";
        WriteFile(file, fileName);
    }

    [TestMethod]
    public void Generate_Excel_Invoices_NoCustomization()
    {
        ExcelGenerator awesomeExcel = new();
        List<Invoice> invoices = GetInvoices();

        MemoryStream file = awesomeExcel.Generate(invoices);

        string fileName = nameof(Generate_Excel_Invoices_NoCustomization) + ".xlsx";
        WriteFile(file, fileName);
    }

    [TestMethod]
    public void Generate_Excel_Invoices()
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

        string fileName = nameof(Generate_Excel_Invoices) + "invoices.xlsx";
        WriteFile(file, fileName);
    }

    private List<Person> GetActors()
    {
        return new List<Person>
        {
            { new() { Name =  "Caroline", Surname = "Aaron", BirthDate = DateTime.Parse("1952-08-07") } },
            { new() { Name =  "Victor", Surname = "Aaron", BirthDate = DateTime.Parse("1956-09-11") } },
            { new() { Name =  "Diego", Surname = "Abatantuono", BirthDate = DateTime.Parse("1955-05-20") } },
            { new() { Name =  "Andrew", Surname = "Abeita", BirthDate = DateTime.Parse("1981-07-11") } },
            { new() { Name =  "Jon", Surname = "Abrahams", BirthDate = DateTime.Parse("1977-10-29") } },
            { new() { Name =  "Stefano", Surname = "Accorsi", BirthDate = DateTime.Parse("1971-03-02") } },
            { new() { Name =  "Dean", Surname = "Acheson", BirthDate = DateTime.Parse("1893-04-11") } },
            { new() { Name =  "Josh", Surname = "Ackerman", BirthDate = DateTime.Parse("1977-03-23") } },
            { new() { Name =  "Joss", Surname = "Ackland", BirthDate = DateTime.Parse("1928-02-29") } },
            { new() { Name =  "Jay", Surname = "Acovone", BirthDate = DateTime.Parse("1955-08-20") } },
            { new() { Name =  "Deb", Surname = "Adair", BirthDate = DateTime.Parse("1966-04-22") } },
            { new() { Name =  "Enid-Raye", Surname = "Adams", BirthDate = DateTime.Parse("1973-06-16") } },
            { new() { Name =  "Jacob", Surname = "Adams", BirthDate = DateTime.Parse("1975-07-04") } },
            { new() { Name =  "Mario", Surname = "Adorf", BirthDate = DateTime.Parse("1930-09-08") } },
            { new() { Name =  "Ben", Surname = "Affleck", BirthDate = DateTime.Parse("1972-08-15") } },
            { new() { Name =  "Casey", Surname = "Affleck", BirthDate = DateTime.Parse("1975-08-12") } },
            { new() { Name =  "Spiro", Surname = "Agnew", BirthDate = DateTime.Parse("1918-11-09") } },
            { new() { Name =  "Antonio", Surname = "Agri", BirthDate = DateTime.Parse("1932-05-05") } },
            { new() { Name =  "Jenny", Surname = "Agutter", BirthDate = DateTime.Parse("1952-12-20") } },
            { new() { Name =  "Betsy", Surname = "Aidem", BirthDate = DateTime.Parse("1957-10-28") } },
            { new() { Name =  "Liam", Surname = "Aiken", BirthDate = DateTime.Parse("1990-01-07") } },
            { new() { Name =  "Troy", Surname = "Aikman", BirthDate = DateTime.Parse("1966-11-21") } },
            { new() { Name =  "Kacey", Surname = "Ainsworth", BirthDate = DateTime.Parse("1970-10-19") } },
            { new() { Name =  "Holly", Surname = "Aird", BirthDate = DateTime.Parse("1969-05-18") } },
            { new() { Name =  "Lucy", Surname = "Akhurst", BirthDate = DateTime.Parse("1975-11-18") } },
            { new() { Name =  "Amy", Surname = "Alcott", BirthDate = DateTime.Parse("1956-02-22") } },
            { new() { Name =  "Alan", Surname = "Alda", BirthDate = DateTime.Parse("1936-01-28") } },
            { new() { Name =  "Tom", Surname = "Aldredge", BirthDate = DateTime.Parse("1928-02-28") } },
            { new() { Name =  "Buzz", Surname = "Aldrin", BirthDate = DateTime.Parse("1930-01-20") } },
            { new() { Name =  "Henry", Surname = "Alessandroni", BirthDate = DateTime.Parse("1959-05-26") } },
            { new() { Name =  "Art", Surname = "Alexakis", BirthDate = DateTime.Parse("1962-04-12") } },
            { new() { Name =  "Jane", Surname = "Alexander", BirthDate = DateTime.Parse("1939-10-28") } },
            { new() { Name =  "Jason", Surname = "Alexander", BirthDate = DateTime.Parse("1959-09-23") } },
            { new() { Name =  "Khandi", Surname = "Alexander", BirthDate = DateTime.Parse("1957-09-04") } },
            { new() { Name =  "Adam", Surname = "Alexi-Malle", BirthDate = DateTime.Parse("1964-09-24") } },
            { new() { Name =  "Hans", Surname = "Alfredson", BirthDate = DateTime.Parse("1931-06-28") } },
            { new() { Name =  "Mary", Surname = "Alice", BirthDate = DateTime.Parse("1941-12-03") } },
            { new() { Name =  "Debbie", Surname = "Allen", BirthDate = DateTime.Parse("1950-01-16") } }
        };
    }

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

    private void WriteFile(MemoryStream file, string fileName)
    {
        string directory = Path.Combine(Environment.CurrentDirectory, "tests-output", nameof(Generate_Excel_AwesomeExcel));
        Directory.CreateDirectory(directory);

        string filePath = Path.Combine(directory, fileName);
        byte[] fileBytes = file.ToArray();
        File.WriteAllBytes(filePath, fileBytes);
    }
}

internal class Person
{
    public string Name { get; set; }
    public string Surname { get; set; }
    public DateTime? BirthDate { get; set; }
}

internal class Invoice
{
    public DateTime CreationDate { get; set; }
    public double Amount { get; set; }
}