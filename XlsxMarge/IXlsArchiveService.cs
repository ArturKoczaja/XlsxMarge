using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    public interface IXlsArchiveService
    {
        IEnumerable<XlsFileToMerge> UnzipFiles(IEnumerable<DocumentItem> documentItems);
        Stream ZipFilesAsStream(object obj);
    }
}