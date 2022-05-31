using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    public class SheetEntry
    {
        public SheetEntry()
        {
        }
        public string FileName { get; set; }
        public byte [] sheetBytes { get; set; }
        public byte [] stringsBytes { get; set; }
    }
}