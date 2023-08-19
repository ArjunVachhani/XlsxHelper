namespace XlsxHelper;

public sealed class Worksheet
{
    public Worksheet(string name, WorksheetReader worksheetReader)
    {
        Name = name;
        WorksheetReader = worksheetReader;
    }

    public string Name { get; private set; }
    public WorksheetReader WorksheetReader { get; private set; }
}
