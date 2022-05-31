using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    public class SheetEntry
    {
        public SheetEntry()
        {
        }

        public SheetEntry(IEnumerable<byte[]> data)
        {
            Data = data;
        }

        public string FileName { get; set; }
        public IEnumerable<byte[]> Data { get; set; }
        
        public byte [] sheetBytes { get; set; }

        public byte [] stringsBytes { get; set; }
    }
}