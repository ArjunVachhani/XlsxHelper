namespace XlsxHelper;

public readonly struct ColumnName
{
    private const char zeroChar = (char)0;

    private readonly char _char1;
    private readonly char _char2;
    private readonly char _char3;

    internal ColumnName(char char1, char char2, char char3)
    {
        //todo check column name range A-Z, uppercase a-z
        _char1 = char1;
        _char2 = char2;
        _char3 = char3;
    }

    [Obsolete("Use parameterized constructor", true)]
    public ColumnName()
    {
        throw new NotImplementedException();
    }

    public static implicit operator ColumnName(string column)
    {
        if (column == null)
            throw new ArgumentNullException(nameof(column));

        if (column.Length <= 0 || column.Length > 3)
            throw new ArgumentException("Column length should be more than 0 and less than 4.");

        var char1 = column[0];
        var char2 = column.Length >= 2 ? column[1] : zeroChar;
        var char3 = column.Length == 3 ? column[2] : zeroChar;

        return new ColumnName(char1, char2, char3);
    }

    public static bool operator ==(ColumnName x, string y)
    {
        return StringEquals(x, y);
    }

    public static bool operator !=(ColumnName x, string y)
    {
        return !StringEquals(x, y);
    }

    public static bool operator ==(ColumnName x, ColumnName y)
    {
        return ColumnEquals(x, y);
    }

    public static bool operator !=(ColumnName x, ColumnName y)
    {
        return !ColumnEquals(x, y);
    }

    public static ColumnName GetColumnName(string cellNumber)
    {
        if (cellNumber == null)
            throw new ArgumentNullException(nameof(cellNumber));

        if (cellNumber.Length < 2)
            throw new ArgumentException("cellNumber length should be at least 2 character long.");

        var char1 = IsValidColumnChar(cellNumber[0]) ? cellNumber[0] : throw new XlsxHelperException($"Invalid cellNumber {cellNumber}.");
        var char2 = IsValidColumnChar(cellNumber[1]) ? cellNumber[1] : zeroChar;
        var char3 = cellNumber.Length >= 3 && IsValidColumnChar(cellNumber[2]) ? cellNumber[2] : zeroChar;

        return new ColumnName(char1, char2, char3);
    }

    private static bool IsValidColumnChar(char ch)
    {
        return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
    }

    private static bool StringEquals(ColumnName x, string y)
    {
        if (y == null)
            throw new ArgumentNullException(nameof(y));

        if (y.Length < 0 || y.Length > 3)
            throw new ArgumentException("y length should be more than 0 and less than 4.");

        var char1 = y[0];
        var char2 = y.Length >= 2 ? y[1] : zeroChar;
        var char3 = y.Length == 3 ? y[2] : zeroChar;

        return x._char1 == char1 && x._char2 == char2 && x._char3 == char3;
    }

    private static bool ColumnEquals(ColumnName x, ColumnName y)
    {
        return x._char1 == y._char1 && x._char2 == y._char2 && x._char3 == y._char3;
    }

    public override bool Equals(object? obj)
    {
        if (obj is null)
            return false;
        if (obj is ColumnName column)
            return this == column;
        if (obj is string str)
            return StringEquals(this, str);
        return false;
    }

    public override int GetHashCode()
    {
        long l = 29034;
        l = l << 16;
        l = l | _char1;
        l = l << 16;
        l = l | _char2;
        l = l << 16;
        l = l | _char3;
        return l.GetHashCode();
    }

    public override string ToString()
    {
        if (_char2 == zeroChar && _char3 == zeroChar)
            return _char1.ToString();
        else if (_char3 == zeroChar)
            return $"{_char1}{_char2}";
        else
            return $"{_char1}{_char2}{_char3}";
    }
}
