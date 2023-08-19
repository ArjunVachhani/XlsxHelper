namespace XlsxHelper;

internal interface ISharedStringList : IDisposable
{
    public int Count { get; }
    public void Add(string item);
    public string this[int index] { get; }
}
