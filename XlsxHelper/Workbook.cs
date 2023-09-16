using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace XlsxHelper;

public sealed class Workbook : IDisposable
{
    private readonly ZipArchive _archive;
    private readonly SharedStringLookup _sharedStringLookup;
    private Worksheet[] _workSheets;

    public IEnumerable<Worksheet> Worksheets => _workSheets;
    internal Workbook(ZipArchive archive)
    {
        _archive = archive;
        var workbookZipEntry = _archive.GetEntry("xl/workbook.xml") ?? throw new XlsxHelperException("workbook file not found.");
        var workbookRelationZipEntry = _archive.GetEntry("xl/_rels/workbook.xml.rels") ?? throw new XlsxHelperException("workbook relationships file not found.");
        var sheets = GetSheetNames(workbookZipEntry.Open());
        var sheetRelations = GetSheetRelations(workbookRelationZipEntry.Open());
        var sharedStringTarget = sheetRelations.Values.FirstOrDefault(x => x.Type == Constants.SharedStringRelationshipType).Target ?? "sharedStrings.xml";
        var sharedStringEntry = _archive.GetEntry($"xl/{sharedStringTarget}");
        _sharedStringLookup = new SharedStringLookup(sharedStringEntry?.Open() ?? new MemoryStream(), (sharedStringEntry?.Length ?? 0) > 20_000_000);

        _workSheets = sheets.Where(x => sheetRelations.ContainsKey(x.RelationId))
            .Select(x => new { Sheet = x, ZipEntry = _archive.GetEntry($"xl/{sheetRelations[x.RelationId].Target}") ?? throw new XlsxHelperException($"zip entry not found for {x.SheetName}.") })
            .Select(x => new Worksheet(x.Sheet.SheetName, new WorksheetReader(x.ZipEntry!.Open(), _sharedStringLookup)))
            .ToArray();
    }

    private static List<(string SheetName, string RelationId)> GetSheetNames(Stream stream)
    {
        using (stream)
        {
            var relationshipNamespace = XmlNamespaces.RelationshipsOpenXmlFormat;
            using var reader = XmlReader.Create(stream, new XmlReaderSettings() { IgnoreWhitespace = true, IgnoreComments = true });
            var sheetNamesWithrId = new List<(string sheetName, string rId)>();
            var stack = new Stack<string>();
            while (reader.Read() && !reader.EOF)
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "workbook" && stack.Count == 0)
                {
                    relationshipNamespace = reader.GetAttribute("xmlns:r") == XmlNamespaces.RelationshipsOclc ? XmlNamespaces.RelationshipsOclc : XmlNamespaces.RelationshipsOpenXmlFormat;
                }
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "sheet" && stack.Count == 2 && stack.Peek() == "sheets")
                {
                    var sheetname = reader.GetAttribute("name");
                    var rId = reader.GetAttribute("id", relationshipNamespace);
                    if (!string.IsNullOrWhiteSpace(sheetname) && !string.IsNullOrWhiteSpace(rId))
                        sheetNamesWithrId.Add((sheetname, rId));
                }

                if (!reader.IsEmptyElement && reader.NodeType == XmlNodeType.Element)
                    stack.Push(reader.Name);

                if (reader.NodeType == XmlNodeType.EndElement)
                    stack.Pop();
            }
            return sheetNamesWithrId;
        }
    }

    private Dictionary<string, (string Target, string Type)> GetSheetRelations(Stream stream)
    {
        using (stream)
        {
            using var reader = XmlReader.Create(stream, new XmlReaderSettings() { IgnoreWhitespace = true, IgnoreComments = true });
            Dictionary<string, (string target, string type)> sheetNamesWithrId = new Dictionary<string, (string target, string type)>();
            Stack<string> stack = new Stack<string>();
            while (reader.Read() && !reader.EOF)
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "Relationship" && stack.Count == 1 && stack.Peek() == "Relationships")
                {
                    var target = reader.GetAttribute("Target");
                    var id = reader.GetAttribute("Id");
                    var type = reader.GetAttribute("Type");
                    if (!string.IsNullOrWhiteSpace(target) && !string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(type))
                        sheetNamesWithrId.Add(id, (target, type));
                }

                if (!reader.IsEmptyElement && reader.NodeType == XmlNodeType.Element)
                    stack.Push(reader.Name);

                if (reader.NodeType == XmlNodeType.EndElement)
                    stack.Pop();
            }
            return sheetNamesWithrId;
        }
    }

    public void Dispose()
    {
        _sharedStringLookup.Dispose();
        foreach (var worksheet in _workSheets)
            worksheet.WorksheetReader.Dispose();
        _archive.Dispose();
    }
}
