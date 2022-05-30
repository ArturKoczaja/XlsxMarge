using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    public interface IXlsxFileMergingService
    {
        Stream MergeFiles(IEnumerable<SheetEntry> files, string inputFilePath, string outputFilePath);
    }
}