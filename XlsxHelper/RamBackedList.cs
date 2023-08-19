namespace XlsxHelper
{
    internal class RamBackedList : List<string>, ISharedStringList
    {
        public void Dispose()
        {
            Clear();
        }
    }
}
