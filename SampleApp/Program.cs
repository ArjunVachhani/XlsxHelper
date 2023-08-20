using XlsxHelper;

namespace SampleApp;

internal class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
    }

    static async Task Sample1(string filePath)
    {
        using (var workbook = XlsxReader.OpenWorkbook(filePath))
        {
            foreach (var worksheet in workbook.Worksheets)//read all worksheets
            {
                Console.WriteLine($"Worksheet {worksheet.Name}"); //get name of worksheet
                using var worksheetReader = worksheet.WorksheetReader; //get WorksheetReader from worksheet
                await foreach (var row in worksheetReader) //read row from worksheetreader
                {
                    Console.WriteLine($"Content of row {row.RowNumber}"); //display current row number
                    foreach (var cell in row.Cells) // Display all cell content
                    {
                        Console.WriteLine($"[{cell.CellValue} at ({cell.ColumnName}{row.RowNumber})]");
                    }
                    Console.WriteLine($"Content of row {row.RowNumber} ends.");
                }
            }
        }
    }

    static async Task Sample2(string filePath)
    {
        using (var workbook = XlsxReader.OpenWorkbook(filePath))
        {
            var worksheet = workbook.Worksheets.First();
            bool headerRow = true;
            Dictionary<string, ColumnName> headerLooklup = null;
            await foreach (var row in worksheet.WorksheetReader)
            {
                if (headerRow)
                {
                    headerLooklup = ReadHeader(row);
                    headerRow = false;
                    continue;
                }
                var student = new Student();
                student.FirstName = row[headerLooklup[nameof(Student.FirstName)]].CellValue;
                student.LastName = row[headerLooklup[nameof(Student.LastName)]].CellValue;
                student.Grade = row[headerLooklup[nameof(student.Grade)]].CellValue;
                student.Marks = new Marks
                {
                    Biology = row[headerLooklup[nameof(Marks.Biology)]].GetInt32(),
                    Chemistry = row[headerLooklup[nameof(Marks.Chemistry)]].GetInt32(),
                    Mathematics = row[headerLooklup[nameof(Marks.Mathematics)]].GetInt32(),
                    Physics = row[headerLooklup[nameof(Marks.Physics)]].GetInt32()
                };

                //Process model. 
                //yield return student;
            }

            static Dictionary<string, ColumnName> ReadHeader(Row row)
            {
                Dictionary<string, ColumnName> headerLooklup = new Dictionary<string, ColumnName>();
                foreach (var cell in row.Cells)
                {
                    headerLooklup.Add(cell.CellValue, cell.ColumnName);
                }
                return headerLooklup;
            }
        }
    }
}