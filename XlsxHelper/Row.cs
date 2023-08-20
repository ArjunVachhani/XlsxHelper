namespace XlsxHelper;

public readonly struct Row
{
    private readonly int _rowNumber;
    private readonly Cell[] _cells;

    internal Row(int rowNumber, Cell[] cells)
    {
        _rowNumber = rowNumber;
        _cells = cells;
    }

    [Obsolete("Use parameterized constructor", true)]
    public Row()
    {
        throw new NotImplementedException();
    }

    public int RowNumber => _rowNumber;
    public Cell[] Cells => _cells;
    public Cell? this[string columnName]
    {
        get
        {
            var cell = _cells.FirstOrDefault(x => x.ColumnName == columnName);
            return cell.ColumnName == columnName ? cell : null;
        }
    }
}
