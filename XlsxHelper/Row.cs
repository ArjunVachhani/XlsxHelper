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
            for (int i = 0; i < _cells.Length; i++)
            {
                if (_cells[i].ColumnName == columnName)
                    return _cells[i];
            }

            throw new XlsxHelperException("Cell value does not exist");
        }
    }

    public Cell this[ColumnName columnName]
    {
        get
        {
            for (int i = 0; i < _cells.Length; i++)
            {
                if (_cells[i].ColumnName == columnName)
                    return _cells[i];
            }

            throw new XlsxHelperException("Cell value does not exist");
        }
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
