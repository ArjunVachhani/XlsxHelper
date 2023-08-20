using System.Text;
using System.Xml;

namespace XlsxHelper;

public sealed class WorksheetReader : IAsyncEnumerable<Row>, IDisposable
{
    private readonly XmlReader _reader;
    private readonly ISharedStringLookup _sharedStringLookup;
    private readonly Stack<string> _nodeHierarchy = new Stack<string>();

    private Row? _currentRow;

    //cleared after every cell processing is completed;
    private readonly List<Cell> _tempCells = new List<Cell>();
    private readonly StringBuilder _tempCellTextStringBuilder = new StringBuilder();

    internal WorksheetReader(Stream stream, ISharedStringLookup sharedStringLookup)
    {
        _reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreComments = true, Async = true });
        _sharedStringLookup = sharedStringLookup;
    }

    public void Dispose()
    {
        _reader.Dispose();
        //dont dispose _sharedStringLookup. it might be used by other WorksheetReader
    }

    public IAsyncEnumerator<Row> GetAsyncEnumerator(CancellationToken cancellationToken = default)
    {
        return new WorksheetRowAsyncEnumerator(this);
    }

    private async ValueTask<bool> MoveNextAsync()
    {
        var rowNumber = 0;
        ColumnName? columnName = null;
        CellValueType? cellValueType = null;
        string? cellValueText = null;
        bool hasMultipleTextForCell = false;
        while (await _reader.ReadAsync() && !_reader.EOF)
        {
            if (_reader.NodeType == XmlNodeType.Element && IsRowElementNode())
            {
                _currentRow = null;
                rowNumber = int.Parse(_reader.GetAttribute("r")!);
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
                    _currentRow = new Row(rowNumber, _tempCells.ToArray());
                    _tempCells.Clear();
                    return true;
                }

                if (IsCellElementNode())
                {
                    if (columnName != null)
                    {
                        var cellText = hasMultipleTextForCell ? _tempCellTextStringBuilder.ToString() : cellValueText;
                        var cell = new Cell(columnName.Value, cellText);
                        _tempCells.Add(cell);
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

    private struct WorksheetRowAsyncEnumerator : IAsyncEnumerator<Row>
    {
        private readonly WorksheetReader _worksheetReader;
        public WorksheetRowAsyncEnumerator(WorksheetReader worksheetReader)
        {
            _worksheetReader = worksheetReader;
        }

        public Row Current => _worksheetReader._currentRow!.Value;

        public ValueTask DisposeAsync()
        {
            _worksheetReader.Dispose();
            return ValueTask.CompletedTask;
        }

        public ValueTask<bool> MoveNextAsync()
        {
            return _worksheetReader.MoveNextAsync();
        }
    }
}
