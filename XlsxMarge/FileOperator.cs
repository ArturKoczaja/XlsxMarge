using System.IO;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    public class FileOperator
    {
        public void ReadFileToByteArrays(ZipFile zipFile, ref MemoryStream sheetStream, ref MemoryStream stringsStream)
        {
            foreach (ZipEntry zipEntry in zipFile)
            {
                if (!zipEntry.IsFile)
                {
                    continue; // Ignore directories
                }

                string entryFileName = zipEntry.Name;
                if (entryFileName == FilesName.SheetName)
                {
                    sheetStream = ZipEntryToStream(zipFile, zipEntry);
                    sheetStream.Position = 0;
                }

                if (entryFileName == FilesName.SharedStringsName)
                {
                    stringsStream = ZipEntryToStream(zipFile, zipEntry);
                    stringsStream.Position = 0;
                }
            }
        }

        public void RemoveSheetAndStringsFiles(ZipFile zipFile)
        {
            zipFile.BeginUpdate();

            foreach (ZipEntry zipEntry in zipFile)
            {
                if (!zipEntry.IsFile)
                {
                    continue; // Ignore directories
                }

                var entryFileName = zipEntry.Name;
                if (entryFileName == FilesName.SheetName)
                {
                    zipFile.Delete(zipEntry);
                }

                if (entryFileName == FilesName.SharedStringsName)
                {
                    zipFile.Delete(zipEntry);
                }
            }

            zipFile.CommitUpdate();
        }

        private MemoryStream ZipEntryToStream(ZipFile zipFile, ZipEntry zipEntry)
        {
            byte[] buffer = new byte[4096];     // 4K is optimum
            Stream zipStream = zipFile.GetInputStream(zipEntry);
            MemoryStream streamWriter = new MemoryStream();
            StreamUtils.Copy(zipStream, streamWriter, buffer);

            return streamWriter;
        }
    }
}