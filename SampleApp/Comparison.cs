using ExcelDataReader;
using LightWeightExcelReader;
using System.Diagnostics;
using XlsxHelper;

namespace SampleApp;

internal class Comparison
{
    public static void XlsxHelper(string filePath)
    {
        var process = Process.GetCurrentProcess();
        using var workbook = XlsxReader.OpenWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();
        Console.WriteLine($"Reading {worksheet.Name}");
        Stopwatch sw = Stopwatch.StartNew();
        var i = 0;
        foreach (var row in worksheet.WorksheetReader)
        {
            if (i++ == 0)
            {
                var ttfrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read first row
                process.Refresh();
                Console.WriteLine($"Time to read first row : {ttfrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
            }
        }
        var ttlrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read last row
        process.Refresh();
        Console.WriteLine($"Time to read all row : {ttlrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
    }

    public static void ExcelDataReader(string filePath)
    {
        var process = Process.GetCurrentProcess();
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            var i = 0;
            Stopwatch sw = Stopwatch.StartNew();
            var ttfrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read first row
            process.Refresh();
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    while (reader.Read())
                    {
                        if (i++ == 0)
                        {
                            Console.WriteLine($"Time to read first row : {ttfrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
                        }
                    }
                } while (reader.NextResult());

                var ttlrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read last row
                process.Refresh();
                Console.WriteLine($"Time to read all row : {ttlrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
            }
        }
    }

    public static void ExcelDataReaderAsDataset(string filePath)
    {
        var process = Process.GetCurrentProcess();
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            Stopwatch sw = Stopwatch.StartNew();
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                reader.AsDataSet();
                var ttlrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read last row
                process.Refresh();
                Console.WriteLine($"Time to read all row : {ttlrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
            }
        }
    }

    public static void LightweightExcelReader(string filePath)
    {
        var process = Process.GetCurrentProcess();
        var excelReader = new ExcelReader(filePath);
        Stopwatch sw = Stopwatch.StartNew();
        var sheetReader = excelReader[0];
        var i = 0;
        while (sheetReader.ReadNext())
        {
            if (i++ == 0)
            {
                var ttfrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read first row
                process.Refresh();
                Console.WriteLine($"Time to read first row : {ttfrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
            }
        }
        var ttlrElapsedMilliseconds = sw.ElapsedMilliseconds;//time to read last row
        process.Refresh();
        Console.WriteLine($"Time to read all row : {ttlrElapsedMilliseconds}ms. Memory Usage : {process.WorkingSet64} bytes");
    }
}
