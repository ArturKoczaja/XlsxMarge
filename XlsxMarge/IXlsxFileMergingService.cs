using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    public interface IXlsxFileMergingService
    {
        byte[] MergeFiles(List<byte[]> inputFilesBytes);
    }
}