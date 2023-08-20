namespace XlsxHelper;

public readonly struct Cell
{
    private readonly ColumnName _columnName;
    private readonly string? _value;

    internal Cell(ColumnName column, string? value)
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
    public string? CellValue => _value;

    public bool TryGetInt32(out int i)
    {
        return int.TryParse(CellValue, out i);
    }

    public bool TryGetInt64(out long l)
    {
        return long.TryParse(CellValue, out l);
    }

    public bool TryGetFloat(out float f)
    {
        return float.TryParse(CellValue, out f);
    }

    public bool TryGetDouble(out double d)
    {
        return double.TryParse(CellValue, out d);
    }

    public bool TryGetDecimal(out decimal d)
    {
        return decimal.TryParse(CellValue, out d);
    }

    public bool TryGetDateTime(out DateTime dt)
    {
        if (double.TryParse(CellValue, out var d))
        {
            try
            {
                dt = DateTime.FromOADate(d);
                return true;
            }
            catch { }
        }
        dt = DateTime.MinValue;
        return false;
    }

    public int GetInt32()
    {
        return int.Parse(CellValue!);
    }

    public long GetInt64()
    {
        return long.Parse(CellValue!);
    }

    public float GetFloat()
    {
        return float.Parse(CellValue!);
    }

    public double GetDouble()
    {
        return double.Parse(CellValue!);
    }

    public decimal GetDecimal()
    {
        return decimal.Parse(CellValue!);
    }

    public DateTime GetDateTime()
    {
        var d = double.Parse(CellValue!);
        return DateTime.FromOADate(d);
    }
}