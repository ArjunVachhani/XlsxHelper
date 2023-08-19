using System.IO.Compression;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("XlsxHelper.Test")]
namespace XlsxHelper;

public static class XlsxReader
{
    public static Workbook OpenWorkbook(string filePath)
    {
        return OpenWorkbook(File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read));
    }

    public static Workbook OpenWorkbook(FileStream fileStream, bool leaveOpen = false)
    {
        return new Workbook(new ZipArchive(fileStream, ZipArchiveMode.Read, leaveOpen));
    }
}
