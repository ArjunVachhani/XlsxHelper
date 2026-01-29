namespace XlsxHelper;

public readonly struct Row : IEnumerable<Cell>
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

    public Cell this[string columnName]
    {
        get
        {
            if (!TryGetCellValue(columnName, out Cell cell))
                throw new XlsxHelperException("Cell value does not exist");
            return cell;
        }
    }

    public Cell this[ColumnName columnName]
    {
        get
        {
            if (!TryGetCellValue(columnName, out var cell))
                throw new XlsxHelperException("Cell value does not exist");
            return cell;
        }
    }

    public bool TryGetCellValue(string columnName, out Cell value)
    {
        for (int i = 0; i < _cells.Length; i++)
        {
            if (_cells[i].ColumnName == columnName)
            {
                value = _cells[i];
                return true;
            }
        }

        value = default;
        return false;
    }

    public bool TryGetCellValue(ColumnName columnName, out Cell value)
    {
        for (int i = 0; i < _cells.Length; i++)
        {
            if (_cells[i].ColumnName == columnName)
            {
                value = _cells[i];
                return true;
            }
        }

        value = default;
        return false;
    }

    public IEnumerator<Cell> GetEnumerator()
    {
        return ((IEnumerable<Cell>)_cells).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return _cells.GetEnumerator();
    }
}
