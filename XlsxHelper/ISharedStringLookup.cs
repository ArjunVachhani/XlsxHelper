namespace XlsxHelper;

internal interface ISharedStringLookup : IDisposable
{
    string GetValue(int position);
}
