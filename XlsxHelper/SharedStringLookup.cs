using System.Text;
using System.Xml;

namespace XlsxHelper;

internal class SharedStringLookup
{
    private readonly XmlReader _reader;
    private readonly ISharedStringList _values;
    private readonly Stack<string> _nodeHierarchy = new Stack<string>();

    //cleared after every si element is processed
    private readonly StringBuilder _tempCellTextStringBuilder = new StringBuilder();
    public SharedStringLookup(Stream xmlStream, bool useFilebackedList)
    {
        _reader = XmlReader.Create(xmlStream, new XmlReaderSettings() { IgnoreComments = true, Async = true });
        _values = useFilebackedList ? new FileBackedList() : new RamBackedList();
    }

    public void Dispose()
    {
        _reader.Dispose();
    }

    public string GetValue(int index)
    {
        if (index > _values.Count - 1)
        {
           var cellText = ReadTill(index);

            if (index > _values.Count - 1)
                throw new XlsxHelperException("Invalid shared string lookup position.");

            return cellText!;
        }

        return _values[index];
    }

    private string? ReadTill(int index)
    {
        string? lastCellText = null;
        bool hasMultipleTextForCell = false;
        string? cellValueText = null;
        while (index >= _values.Count && _reader.Read() && !_reader.EOF)
        {
            if (IsSiTextNode() || IsSiRichTextNode())
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

                if (IsSiElementNode())
                {
                    var cellText = hasMultipleTextForCell ? _tempCellTextStringBuilder.ToString() : cellValueText;
                    _values.Add(cellText!);
                    lastCellText = cellText;
                    hasMultipleTextForCell = false;
                    cellValueText = null;
                    _tempCellTextStringBuilder.Clear();
                }
            }
        }
        return lastCellText;
    }

    private bool IsSiElementNode() => _reader.Name == "si" && _nodeHierarchy.Count == 1 && _nodeHierarchy.Peek() == "sst";
    private bool IsSiTextNode() => (_reader.NodeType == XmlNodeType.Text || _reader.NodeType == XmlNodeType.Whitespace || _reader.NodeType == XmlNodeType.SignificantWhitespace) && _nodeHierarchy.Count == 3 && _nodeHierarchy.Peek() == "t";
    private bool IsSiRichTextNode() => (_reader.NodeType == XmlNodeType.Text || _reader.NodeType == XmlNodeType.Whitespace || _reader.NodeType == XmlNodeType.SignificantWhitespace) && _nodeHierarchy.Count == 4 && _nodeHierarchy.Peek() == "t";

}
