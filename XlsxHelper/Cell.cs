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
    public string? Value => _value;

    public bool TryGetInt32(out int i)
    {
        return int.TryParse(Value, out i);
    }

    public bool TryGetInt64(out long l)
    {
        return long.TryParse(Value, out l);
    }

    public bool TryGetFloat(out float f)
    {
        return float.TryParse(Value, out f);
    }

    public bool TryGetDouble(out double d)
    {
        return double.TryParse(Value, out d);
    }

    public bool TryGetDecimal(out decimal d)
    {
        return decimal.TryParse(Value, out d);
    }

    public bool TryGetDateTime(out DateTime dt)
    {
        if (double.TryParse(Value, out var d))
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
        return int.Parse(Value!);
    }

    public long GetInt64()
    {
        return long.Parse(Value!);
    }

    public float GetFloat()
    {
        return float.Parse(Value!);
    }

    public double GetDouble()
    {
        return double.Parse(Value!);
    }

    public decimal GetDecimal()
    {
        return decimal.Parse(Value!);
    }

    public DateTime GetDateTime()
    {
        var d = double.Parse(Value!);
        return DateTime.FromOADate(d);
    }
}