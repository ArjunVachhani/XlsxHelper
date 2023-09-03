# XlsxHelper
![](https://github.com/ArjunVachhani/XlsxHelper/workflows/.NET/badge.svg?branch=main)

XlsxHelper has been crafted with the primary intention of efficiently parsing extensive Excel files. This library is designed to be lightweight, ensuring that it doesn't load the entire dataset into memory all at once. Instead, it adopts a sequential approach, fetching and returning one row per iteration. As a result of this methodology, the memory overhead is minimized.

### When to use XlsxHelper
- You need to process large xlsx file.
- You want to read content with very little RAM usage.
- You want full control of mapping/parsing fields to Model.

### When to not use XlsxHelper
- You want to read rows in random manner.
- You want to read thing like width of row/column, font size / color etc 
- You want to read xls file format.

### Project Status
XlsxHelper is actively maintained. Please feel free to ask question and raise issues.

### How to get started
Install NuGet package https://www.nuget.org/packages/XlsxHelper/

### Example 1
```csharp
using (var workbook = XlsxReader.OpenWorkbook(filePath))
{
    foreach (var worksheet in workbook.Worksheets)//read all worksheets
    {
        Console.WriteLine($"Worksheet {worksheet.Name}"); //get name of worksheet
        using var worksheetReader = worksheet.WorksheetReader; //get WorksheetReader from worksheet
        foreach (var row in worksheetReader) //read row from worksheetreader
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
```        

### Example 2
```csharp
using (var workbook = XlsxReader.OpenWorkbook(filePath))
{
    var worksheet = workbook.Worksheets.First();
    bool headerRow = true;
    Dictionary<string, ColumnName> headerLooklup = null;
    foreach (var row in worksheet.WorksheetReader)
    {
        if (headerRow)
        {
            headerLooklup = ReadHeader(row); //Read all header names from first row
            headerRow = false;
            continue;
        }
        var student = new Student();
        student.FirstName = row[headerLooklup[nameof(Student.FirstName)]].CellValue; //get cell value 
        student.LastName = row[headerLooklup[nameof(Student.LastName)]].CellValue;
        student.Grade = row[headerLooklup[nameof(student.Grade)]].CellValue;
        student.Marks = new Marks
        {
            Biology = row[headerLooklup[nameof(Marks.Biology)]].GetInt32(),
            Chemistry = row[headerLooklup[nameof(Marks.Chemistry)]].GetInt32(),
            Mathematics = row[headerLooklup[nameof(Marks.Mathematics)]].GetInt32(),
            Physics = row[headerLooklup[nameof(Marks.Physics)]].GetInt32()
        };

        //Process student object. 
        //yield return student;
    }

    static Dictionary<string, ColumnName> ReadHeader(Row row) //read header
    {
        Dictionary<string, ColumnName> headerLooklup = new Dictionary<string, ColumnName>();
        foreach (var cell in row.Cells)
        {
            headerLooklup.Add(cell.CellValue, cell.ColumnName);
        }
        return headerLooklup;
    }
}
```

### How fast and lightweight is it?

|                                               | XlsxHelper | LightweightExcelReader | ExcelDataReader AsDataset | ExcelDataReader |
|-----------------------------------------------|------------|------------------------|---------------------------|-----------------|
| Time to read first row                        | 5ms        | 14ms                   | -                         | -               |
| Time to read all rows(50,000)                 | 3.90 sec   | 7.60 sec               | 13.50 sec                 | 10.10 sec       |
| Memory usage at the time of reading first row | 31.412 MB  | 32.649 MB              | -                         | 42.057 MB       |
| Memory usage at the time of reading last row  | 38.891 MB  | 901.976 MB             | 471.662 MB                | 42.414 MB       |