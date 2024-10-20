using AwesomeExcel;
using AwesomeExcel.Models;
using BenchmarkDotNet.Attributes;

namespace Benchmarks;

public class Generate_Excel_BigDataSets
{
    private bool writefile = false;
    private IEnumerable<Invoice> invoices;

    [GlobalSetup]
    public void Setup()
    {
        invoices = MockData.GetInvoices(5_000);
    }

    [Benchmark]
    public void Generate_Excel_Invoices_5000()
    {
        ExcelGenerator awesomeExcel = new();
        MemoryStream file = Generate(awesomeExcel, invoices);
                                 
        string fileName = nameof(Generate_Excel_Invoices_5000) + ".xlsx";

        WriteFile(file, fileName);
    }

    private static MemoryStream Generate(ExcelGenerator awesomeExcel, IEnumerable<Invoice> invoices)
    {
        // Generate the Excel file with some customizations
        MemoryStream file = awesomeExcel.Generate(invoices, (customization) =>
        {
            // Customize the entire sheet
            customization
                .SetName("Client's invoices")
                .SetFontName("Aptos")
                .SetBordersColor(Color.Black);

            // Customize the header row
            customization
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
            customization.Cells(columns => columns.Random1)
                .SetFillForegroundColor(random1 => random1 > 1000 ? Color.LightGreen : null);
        });
        return file;
    }

    private void WriteFile(MemoryStream file, string fileName)
    {
        if (writefile == false) return;

        string directory = Path.Combine(Environment.CurrentDirectory, "benchmark-excel-output", nameof(Generate_Excel_BigDataSets));
        Directory.CreateDirectory(directory);

        string filePath = Path.Combine(directory, fileName);
        byte[] fileBytes = file.ToArray();
        File.WriteAllBytes(filePath, fileBytes);
    }

    private class Invoice
    {
        public DateTime CreationDate { get; set; }
        public double Amount { get; set; }
        public int Random1 { get; set; }
        public int Random2 { get; set; }
        public int Random3 { get; set; }
        public int Random4 { get; set; }
        public int Random5 { get; set; }
        public int Random6 { get; set; }
        public int Random7 { get; set; }
        public int Random8 { get; set; }
        public int Random9 { get; set; }
        public string Random10 { get; set; }
    }

    private static class MockData
    {
        public static IEnumerable<Invoice> GetRandomInvoices(int count)
        {
            Random random = new();

            for (int i = 0; i < count; i++)
            {
                Invoice invoice = new()
                {
                    Amount = random.NextDouble() * 10000 / 100,
                    CreationDate = new DateTime(random.Next(1950, 2050), random.Next(1, 12), random.Next(1, 28), random.Next(0, 23), random.Next(0, 59), random.Next(0, 59)),
                    Random1 = random.Next(),
                    Random2 = random.Next(),
                    Random3 = random.Next(),
                    Random4 = random.Next(),
                    Random5 = random.Next(),
                    Random6 = random.Next(),
                    Random7 = random.Next(),
                    Random8 = random.Next(),
                    Random9 = random.Next(),
                    Random10 = RandomString(1, 100),
                };
                yield return invoice;
            }
        }

        public static IEnumerable<Invoice> GetInvoices(int count)
        {
            return GetRandomInvoices(count);
        }

        private static string RandomString(int minLength, int maxLength)
        {
            // https://stackoverflow.com/questions/1344221/how-can-i-generate-random-alphanumeric-strings

            Random random = new();

            int length = random.Next(minLength, maxLength);

            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
