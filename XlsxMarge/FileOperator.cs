using System.IO;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    public class FileOperator
    {
        //public void ReadFilesToByteArrays(ZipFile zipFile, ref MemoryStream sheetStream, ref MemoryStream stringsStream)
        //{
        //    foreach (ZipEntry zipEntry in zipFile)
        //    {
        //        if (!zipEntry.IsFile)
        //        {
        //            continue; // Ignore directories
        //        }

        //        string entryFileName = zipEntry.Name;
        //        if (entryFileName == FileNames.SheetName)
        //        {
        //            sheetStream = ZipEntryToStream(zipFile, zipEntry);
        //            sheetStream.Position = 0;
        //        }

        //        if (entryFileName == FileNames.SharedStringsName)
        //        {
        //            stringsStream = ZipEntryToStream(zipFile, zipEntry);
        //            stringsStream.Position = 0;
        //        }
        //    }
        //}

        public MemoryStream ReadFileToStream(ZipFile zipFile, string entryName)
        {
            MemoryStream outStream = null;
            foreach (ZipEntry zipEntry in zipFile)
            {
                if (!zipEntry.IsFile)
                {
                    continue; // Ignore directories
                }

                if (zipEntry.Name == entryName)
                {
                    outStream = ZipEntryToStream(zipFile, zipEntry);
                }
            }
            return outStream;
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
                if (entryFileName == FileNames.SheetName)
                {
                    zipFile.Delete(zipEntry);
                }

                if (entryFileName == FileNames.SharedStringsName)
                {
                    zipFile.Delete(zipEntry);
                }
            }

            zipFile.CommitUpdate();
        }

        private MemoryStream ZipEntryToStream(ZipFile zipFile, ZipEntry zipEntry)
        {
            byte[] buffer = new byte[4096];
            MemoryStream outStream = new MemoryStream();;// 4K is optimum

            using (var zipStream = zipFile.GetInputStream(zipEntry))
            {
                StreamUtils.Copy(zipStream, outStream, buffer);
            }
            outStream.Position = 0;
            return outStream;
        }
    }
}