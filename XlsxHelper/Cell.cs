namespace XlsxHelper;

public readonly struct Cell
{
    //TODO colspan, rowspan, formula
    private readonly ColumnName _columnName;
    private readonly string? _value;

    public Cell(ColumnName column, string? value)
    {
        _columnName = column;
        _value = value;
    }

    [Obsolete("Use parameterized constructor", true)]
    public Cell()
    {
        throw new NotImplementedException();
    }

    public ColumnName ColumnName => _columnName;
    public string? Value => _value;
}