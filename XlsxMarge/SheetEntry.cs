using System.IO;

namespace XlsxMarge
{
    public class SheetEntry
    {
        public string FileName { get; set; }

        public MemoryStream StreamForSheetFile { get; set; }

        public MemoryStream StreamForSharedStringsFile { get; set; }
    }
}