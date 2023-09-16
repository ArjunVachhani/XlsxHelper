namespace XlsxHelper;

public sealed class Worksheet : IEnumerable<Row>
{
    public Worksheet(string name, WorksheetReader worksheetReader)
    {
        Name = name;
        WorksheetReader = worksheetReader;
    }

    public string Name { get; private set; }
    public WorksheetReader WorksheetReader { get; private set; }

    public IEnumerator<Row> GetEnumerator()
    {
        return WorksheetReader.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return WorksheetReader.GetEnumerator();
    }
}
