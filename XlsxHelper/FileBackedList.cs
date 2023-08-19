using System.Buffers;
using System.Text;

namespace XlsxHelper;

public class FileBackedList : ISharedStringList
{
    private readonly FileStream fileStream;
    private readonly SortedSet<RowPointer> pageIndex;
    private readonly byte[] buffer = new byte[128];
    private long appendPosition;
    private int lastIndexedPage = -1;
    private RowPointer? lastRowPointer;

    public FileBackedList()
    {
        fileStream = File.Open(Path.GetTempFileName(), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None | FileShare.Delete);
        appendPosition = fileStream.Position;
        pageIndex = new SortedSet<RowPointer>(RowPointerComparer.Comparer);
    }

    public int Count { get; private set; }

    public string this[int index] => Get(index);

    public void Add(string item)
    {
        if (fileStream.Position != appendPosition)
            fileStream.Seek(appendPosition, SeekOrigin.Begin);

        if ((appendPosition / 4096) > lastIndexedPage)
        {
            lastRowPointer = new RowPointer(appendPosition, Count);
            lastIndexedPage = (int)(appendPosition / 4096);
            pageIndex.Add(lastRowPointer);
        }
        else
        {
            lastRowPointer!.SetLastRowIndex(Count);
        }

        var size = Encoding.UTF8.GetByteCount(item);
        fileStream.Write(BitConverter.GetBytes(size), 0, 4);
        if (size > 0)
        {
            bool rented;
            byte[] bytes;
            if (size <= 128)
            {
                bytes = buffer;
                rented = false;
            }
            else
            {
                bytes = ArrayPool<byte>.Shared.Rent(size);
                rented = true;
            }

            Encoding.UTF8.GetBytes(item, 0, item.Length, bytes, 0);
            fileStream.Write(bytes, 0, size);
            if (rented)
                ArrayPool<byte>.Shared.Return(bytes);
        }

        appendPosition = fileStream.Position;
        Count++;
    }

    public string Get(int index)
    {
        if (index < 0 || index >= Count)
            throw new XlsxHelperException("Index out of range.");

        pageIndex.TryGetValue(new RowPointer(0, index), out var rowPointer);
        if (fileStream.Position / 4096 != rowPointer!.Position / 4096 || fileStream.Position > rowPointer.Position)
            fileStream.Seek(rowPointer.Position, SeekOrigin.Begin);

        var rowIndex = rowPointer.FirstRowIndex;
        int size;
        do
        {
            fileStream.Read(buffer, 0, 4);
            size = BitConverter.ToInt32(buffer, 0);
            if (rowIndex < index && size > 0)
                fileStream.Seek(size, SeekOrigin.Current);
        }
        while (rowIndex++ < index);

        if (size == 0)
            return string.Empty;

        bool rented;
        byte[] bytes;
        if (size <= 128)
        {
            bytes = buffer;
            rented = false;
        }
        else
        {
            bytes = ArrayPool<byte>.Shared.Rent(size);
            rented = true;
        }

        fileStream.Read(bytes, 0, size);
        var str = Encoding.UTF8.GetString(bytes, 0, size);
        if (rented)
            ArrayPool<byte>.Shared.Return(bytes);
        return str;
    }

    public void Dispose()
    {
        fileStream.Dispose();
        pageIndex.Clear();
    }

    private class RowPointer
    {
        private readonly int pageNumber;
        private readonly short offset;
        private readonly int firstRowIndex;
        private int lastRowIndex;

        public RowPointer(long position, int firstRowIndex)
        {
            this.firstRowIndex = firstRowIndex;
            this.lastRowIndex = this.firstRowIndex;
            pageNumber = (int)(position / 4096);
            offset = (short)(position % 4096);
        }

        public long Position => (pageNumber * 4096) + offset;

        // Row Index starts from 0
        public int FirstRowIndex => firstRowIndex;

        public int LastRowIndex => lastRowIndex;

        public void SetLastRowIndex(int lastRowIndex)
        {
            this.lastRowIndex = lastRowIndex;
        }
    }

    private class RowPointerComparer : IComparer<RowPointer>
    {
        public static RowPointerComparer Comparer { get; } = new RowPointerComparer();

        public int Compare(RowPointer? x, RowPointer? y)
        {
            if (x?.FirstRowIndex < y?.FirstRowIndex)
                return -1;
            else if (x?.LastRowIndex > y?.LastRowIndex)
                return 1;
            else
                return 0;
        }
    }
}
