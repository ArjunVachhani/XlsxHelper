# XlsxHelper
lightweight Xlsx data reader

## Example 1
```csharp
        using (var workbook = XlsxReader.OpenWorkbook(path))
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet {worksheet.Name}");
                using var worksheetReader = worksheet.WorksheetReader;
                await foreach (var row in worksheetReader)
                {
                    Console.WriteLine($"Content of row {row.RowNumber}");
                    foreach (var cell in row.Cells)
                    {
                        Console.WriteLine($"[{cell.Value} at ({cell.ColumnName}{row.RowNumber})]");
                    }
                    Console.WriteLine($"Content of row {row.RowNumber} ends.");
                }
            }
        }
```        

## Example 2
```csharp
        using (var workbook = XlsxReader.OpenWorkbook(path))
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                using var worksheetReader = worksheet.WorksheetReader;
                await foreach (var row in worksheetReader)
                {
                    var student = new Student();
                    student.Mark = int.Parse(row.Cells[0].Value);
                    student.Name = row.Cells[1].Value;
                }
            }
        }
        
```