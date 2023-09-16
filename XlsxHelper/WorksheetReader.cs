using System.Text;
using System.Xml;

namespace XlsxHelper;

public sealed class WorksheetReader : IEnumerable<Row>, IDisposable
{
    private readonly XmlReader _reader;
    private readonly SharedStringLookup _sharedStringLookup;
    private readonly Stack<string> _nodeHierarchy = new Stack<string>();

    private int _rowNumber = 0;
    private readonly List<Cell> _cells = new List<Cell>();

    //cleared after every cell processing is completed.
    private readonly StringBuilder _tempCellTextStringBuilder = new StringBuilder();

    internal WorksheetReader(Stream stream, SharedStringLookup sharedStringLookup)
    {
        _reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true });
        _sharedStringLookup = sharedStringLookup;
    }

    public void Dispose()
    {
        _reader.Dispose();
        //dont dispose _sharedStringLookup. it might be used by other WorksheetReader
    }

    public IEnumerator<Row> GetEnumerator()
    {
        return new WorksheetRowEnumerator(this);
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return new WorksheetRowEnumerator(this);
    }

    public bool ReadRow()
    {
        return MoveNext();
    }

    public Cell this[int index]
    {
        get
        {
            return index < _cells.Count ? _cells[index] : throw new XlsxHelperException("Cell value does not exist");
        }
    }

    public Cell this[string columnName]
    {
        get
        {
            for (int i = 0; i < _cells.Count; i++)
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
            for (int i = 0; i < _cells.Count; i++)
            {
                if (_cells[i].ColumnName == columnName)
                    return _cells[i];
            }

            throw new XlsxHelperException("Cell value does not exist");
        }
    }

    private bool MoveNext()
    {
        ColumnName? columnName = null;
        CellValueType? cellValueType = null;
        string? cellValueText = null;
        bool hasMultipleTextForCell = false;
        while (_reader.Read() && !_reader.EOF)
        {
            if (_reader.NodeType == XmlNodeType.Element && IsRowElementNode())
            {
                _cells.Clear();
                _rowNumber = int.Parse(_reader.GetAttribute("r")!);
            }

            if (_reader.NodeType == XmlNodeType.Element && IsCellElementNode())
            {
                columnName = ColumnName.GetColumnName(_reader.GetAttribute("r")!);
                cellValueType = GetCellValueType(_reader.GetAttribute("t")!);
            }

            if (IsCellValueTextNode() && columnName != null && cellValueType != null)
            {
                var text = _reader.Value;
                if (cellValueType == CellValueType.SharedString)
                    text = _sharedStringLookup.GetValue(int.Parse(text));

                cellValueText = text;
            }
            if ((IsInternalStringTextNode() || IsInternalRichTextStringTextNode()) && columnName != null && cellValueType != null)
            {
                //there can be multiple t tag for 1 cell, in that case combile all.
                var text = _reader.Value;
                if (cellValueText != null)
                {
                    if (hasMultipleTextForCell == false)
                        _tempCellTextStringBuilder.Append(cellValueText);

                    hasMultipleTextForCell = true;
                    _tempCellTextStringBuilder.Append(text);
                }
                cellValueText = text;
            }

            if (!_reader.IsEmptyElement && _reader.NodeType == XmlNodeType.Element)
                _nodeHierarchy.Push(_reader.Name);

            if (_reader.NodeType == XmlNodeType.EndElement)
            {
                _nodeHierarchy.Pop();

                if (IsRowElementNode())
                {
                    return true;
                }

                if (IsCellElementNode())
                {
                    if (columnName != null)
                    {
                        var cellText = hasMultipleTextForCell ? _tempCellTextStringBuilder.ToString() : cellValueText;
                        var cell = new Cell(columnName.Value, cellText);
                        _cells.Add(cell);
                    }

                    hasMultipleTextForCell = false;
                    columnName = null;
                    cellValueType = null;
                    cellValueText = null;
                    _tempCellTextStringBuilder.Clear();
                }
            }
        }
        return false;
    }

    private bool IsRowElementNode() => _reader.Name == "row" && _nodeHierarchy.Count == 2 && _nodeHierarchy.Peek() == "sheetData";
    private bool IsCellElementNode() => _reader.Name == "c" && _nodeHierarchy.Count == 3 && _nodeHierarchy.Peek() == "row";
    private bool IsCellValueTextNode() => _reader.NodeType == XmlNodeType.Text && _nodeHierarchy.Count == 5 && _nodeHierarchy.Peek() == "v";
    private bool IsInternalStringTextNode() => (_reader.NodeType == XmlNodeType.Text || _reader.NodeType == XmlNodeType.Whitespace || _reader.NodeType == XmlNodeType.SignificantWhitespace) && _nodeHierarchy.Count == 6 && _nodeHierarchy.Peek() == "t";
    private bool IsInternalRichTextStringTextNode() => (_reader.NodeType == XmlNodeType.Text || _reader.NodeType == XmlNodeType.Whitespace || _reader.NodeType == XmlNodeType.SignificantWhitespace) && _nodeHierarchy.Count == 7 && _nodeHierarchy.Peek() == "t";

    private CellValueType GetCellValueType(string cellValueType)
    {
        return cellValueType switch
        {
            "b" => CellValueType.Boolean,
            "d" => CellValueType.Date,
            "e" => CellValueType.Error,
            "inlineStr" => CellValueType.InlineString,
            "n" => CellValueType.Number,
            "s" => CellValueType.SharedString,
            "str" => CellValueType.Formula,
            _ => CellValueType.Unkown
        };
    }

    private class WorksheetRowEnumerator : IEnumerator<Row>
    {
        private readonly WorksheetReader _worksheetReader;
        public WorksheetRowEnumerator(WorksheetReader worksheetReader)
        {
            _worksheetReader = worksheetReader;
            Current = new Row(-1, new Cell[0]);
        }

        public Row Current { get; private set; }

        object IEnumerator.Current => Current;

        public void Dispose()
        {
            _worksheetReader.Dispose();
        }

        public bool MoveNext()
        {
            var result = _worksheetReader.MoveNext();
            Current = new Row(_worksheetReader._rowNumber, _worksheetReader._cells.ToArray());
            return result;
        }

        public void Reset()
        {
            throw new InvalidOperationException();
        }
    }
}
