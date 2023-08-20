namespace XlsxHelper.Test;

public class WorkbookTest
{
    [Theory]
    [InlineData("empty.xlsx")]
    [InlineData("multipleemptysheets.xlsx")]
    public async Task EmptyXlsx(string fileName)
    {
        var path = Utility.GetXlsxSampleFilePath(fileName);
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

    [Theory]
    [InlineData("invalidfile.xlsx")]
    public void NonZipFile(string fileName)
    {
        var path = Utility.GetXlsxSampleFilePath(fileName);
        Assert.Throws<InvalidDataException>(() => { XlsxReader.OpenWorkbook(path); });
    }

    [Theory]
    [InlineData("missingworkbook.xlsx")]
    [InlineData("missingworkbookrelatioship.xlsx")]
    public void InvalidXlsx(string fileName)
    {
        var path = Utility.GetXlsxSampleFilePath(fileName);
        Assert.Throws<XlsxHelperException>(() => { XlsxReader.OpenWorkbook(path); });
    }

    [Theory]
    [InlineData("multisheet1.xlsx", new[] { "one", "two", "three", "b", "a" })]
    [InlineData("singlesheet.xlsx", new[] { "one" })]
    public async Task IdentifiesWorksheets(string fileName, string[] worksheetNames)
    {
        var path = Utility.GetXlsxSampleFilePath(fileName);
        using var workbook = XlsxReader.OpenWorkbook(path);
        Assert.Equal(worksheetNames.Length, workbook.Worksheets.Count());
        int i = 0;
        foreach (var worksheet in workbook.Worksheets)
        {
            Assert.Equal(worksheetNames[i], worksheet.Name);
            using var worksheetReader = worksheet.WorksheetReader;
            await foreach (var row in worksheetReader)
            {
            }
            i++;
        }
    }

    [Theory]
    [InlineData("verysimple.xlsx")]
    public async Task DisposeRelasesFile(string fileName)
    {
        var path = Utility.GetXlsxSampleFilePath(fileName);
        using (var workbook = XlsxReader.OpenWorkbook(path))
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                using var worksheetReader = worksheet.WorksheetReader;
                await foreach (var row in worksheetReader)
                {
                    //read lock is held
                    Assert.Throws<IOException>(() => File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
                }
            }
        }
        //no lock, open read/write should work
        using var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

    [Theory]
    [InlineData("verysimple.xlsx")]
    public void DisposeRelasesFile2(string fileName)
    {
        var path = Utility.GetXlsxSampleFilePath(fileName);
        using (var workbook = XlsxReader.OpenWorkbook(path))
        {
            //read lock is held
            Assert.Throws<IOException>(() => File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
        }
        //no lock, open read/write should work
        using var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }
}
