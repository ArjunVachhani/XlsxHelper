namespace XlsxHelper.Test;

public class WorksheetReaderTest
{
    [Fact]
    public void SampleFile()
    {
        var path = Utility.GetXlsxSampleFilePath("verysimple.xlsx");
        using var workbook = XlsxReader.OpenWorkbook(path);
        Assert.NotEmpty(workbook.Worksheets);
        foreach (var worksheet in workbook.Worksheets)
        {
            using var worksheetReader = worksheet.WorksheetReader;
            foreach (var row in worksheetReader)
            {
            }
        }
    }

    [Fact]
    public void StyleAndFormattedFile()
    {
        var path = Utility.GetXlsxSampleFilePath("styledworkbook.xlsx");
        var workSheet1Content = new string[][]
        {
            new string[]{ "a1", "multiline line1\nMultiline line2\nMultiline line 3 multi word", "c1", "d1", "e1" },
            new string[]{ "bold", "italic", "bold italic", "bold italic underline" },
            new string[]{ "bg color1", "bg color and font color", "font color", "text size changed" },
            new string[]{ "font changed", "Font + size changed", "<", "&", "'" },
            new string[]{ "“", "<html>", "<script></script>", "<?xml ?> " },
            new string[]{ "multi format" , "\"", " text  ", " t", "t " },
            new string[]{ "करो हाथों को ऊपर कस आवी गयो", "કેમ છો " }
        };
        using var workbook = XlsxReader.OpenWorkbook(path);
        Assert.NotEmpty(workbook.Worksheets);
        var worksheet1 = workbook.Worksheets.First();
        Assert.Equal("text styling", worksheet1.Name);
        var r = 0;
        foreach (var row in worksheet1.WorksheetReader)
        {
            int c = 0;
            for (; c < row.Cells.Length; c++)
            {
                Assert.Equal(workSheet1Content[r][c], row.Cells[c].CellValue);
            }
            Assert.Equal(workSheet1Content[r].Length, c);
            r++;
        }
        Assert.Equal(workSheet1Content.Length, r);

        var workSheet2Content = new string[][]
        {
            new string[]{ "123", "2022", "12"  },
            new string[]{ "123.749273492379", "Mar – 2022", "12.79879" },
            new string[]{ "123.749273492379", "44621", "1232.1" },
            new string[]{ "12313.123123123", "18 mar 22", "123" },
            new string[]{ "13", "200" },
            new string[]{ "0.00129", "200.90909" },
            new string[]{ "999.999999", "8980" },
            new string[]{ "999.999999", "0.508333333333333" },
            new string[]{ "23.3" },
            new string[]{ "1" },
            new string[]{ "2" },
            new string[]{ "2" },
            new string[]{ "-1"},
            new string[]{ "0" },
            new string[]{ "0.5" },
            new string[]{ "0.25" }
        };
        var worksheet2 = workbook.Worksheets.ElementAt(1);
        Assert.Equal("number & date formatting", worksheet2.Name);
        r = 0;
        foreach (var row in worksheet2.WorksheetReader)
        {
            int c = 0;
            for (; c < row.Cells.Length; c++)
            {
                Assert.Equal(workSheet2Content[r][c], row.Cells[c].CellValue);
            }
            Assert.Equal(workSheet2Content[r].Length, c);
            r++;
        }
        Assert.Equal(workSheet2Content.Length, r);
    }
}
