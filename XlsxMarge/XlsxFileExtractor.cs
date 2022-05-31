using System.Collections.Generic;
using System.IO;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    public class XlsxFileExtractor
    {
        private IEnumerable<byte[]> data;

        public List<SheetEntry> UnzipXlsxFiles(List<byte[]> inputBytes)
        {
            var sheets = new List<SheetEntry>();

            foreach (var fileBytes in inputBytes)
            {
                sheets.Add(ExtractSheetFiles(fileBytes));
            }

            return sheets;
        }

        private SheetEntry ExtractSheetFiles(byte [] bytes)
        {
            using (var fileStream = new MemoryStream(bytes))
            {
                using (var zipEntries = new ZipFile(fileStream))
                {
                    int counter = 0;

                    byte[] sheetArray = null;
                    byte[] stringsArray = null;

                    foreach (ZipEntry zipEntry in zipEntries)
                    {
                        if (!zipEntry.IsFile)
                        {
                            continue; // Ignore directories
                        }

                        string entryFileName = zipEntry.Name;
                        if (entryFileName == FilesName.SheetName)
                        {
                            sheetArray = ZipEntryToStream(zipEntries, zipEntry);

                            counter++;
                        }

                        if (entryFileName == FilesName.SharedStringsName)
                        {
                            stringsArray = ZipEntryToStream(zipEntries, zipEntry);
                            counter++;
                        }

                        if (counter == 2)
                        {
                            // if both files read we can stop
                            break;
                        }
                    }

                    if (sheetArray is null || stringsArray is null)
                    {
                        return null;
                    }

                    return new SheetEntry()
                    {
                        sheetBytes = sheetArray,
                        stringsBytes = stringsArray,
                        Data = data
                    };
                }
            }
        }

        private byte [] ZipEntryToStream(ZipFile zf, ZipEntry zipEntry)
        {
            byte[] buffer = new byte[4096];     // 4K is optimum
            Stream zipStream = zf.GetInputStream(zipEntry);

            using MemoryStream streamWriter = new MemoryStream();
            StreamUtils.Copy(zipStream, streamWriter, buffer);
            streamWriter.Position = 0;
            return streamWriter.ToArray();

        }
    }
}