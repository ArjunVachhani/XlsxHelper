namespace XlsxHelper.Test;

public class WorksheetReaderTest
{
    [Fact]
    public async Task SampleFile()
    {
        var path = Utility.GetXlsxSampleFilePath("verysimple.xlsx");
        using var workbook = XlsxReader.OpenWorkbook(path);
        Assert.NotEmpty(workbook.Worksheets);
        foreach (var worksheet in workbook.Worksheets)
        {
            using var worksheetReader = worksheet.WorksheetReader;
            await foreach (var row in worksheetReader)
            {
            }
        }
    }
}
