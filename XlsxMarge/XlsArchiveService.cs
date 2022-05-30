using System;
using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    public class XlsArchiveService : IXlsArchiveService
    {
        public IEnumerable<XlsFileToMerge> UnzipFiles(IEnumerable<DocumentItem> documentItems)
        {
            // Loop through documentItems, unzip all files and return list of XlsFileToMerge (use UnzipFile method).
            // Then this list will be used in XlsxFileMergingService to merge all those files to single one.
            throw new NotImplementedException();
        }

        // Unzip data from documentItem.Data
        private IEnumerable<XlsFileToMerge> UnzipFile(DocumentItem documentItem)
        {
            var xlsFilesToMerge = new List<XlsFileToMerge>();

            foreach (var fileData in documentItem.Data)
            {
                // Check if If SharpZipLib can unzip file in byte[] format
                // 1. If no first of all, byte[] must be converted to Stream then Stream will be unzip
                // 2. Then try to iterate through single unzip file (check ZipEntry or similar class) and try to retrieve sheet1.xml and sharedStrings.xml files
                // 3. Create XlsFileToMerge and add there those two files and file name
                // 4. Add newly created XlsFileToMerge to the list xlsFilesToMerge
            }

            return xlsFilesToMerge;
        }

        public Stream ZipFilesAsStream(object obj)
        {
            throw new NotImplementedException();
        }
    }
}