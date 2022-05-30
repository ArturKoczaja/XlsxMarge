using System.Collections.Generic;
using System.IO;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    public class XlsxFileExtractor
    {
        public List<SheetEntry> UnzipXlsxFiles(string[] inputFiles)
        {
            var sheets = new List<SheetEntry>();

            foreach (var file in inputFiles)
            {
                sheets.Add(ExtractSheetFiles(file));
            }

            return sheets;
        }

        private SheetEntry ExtractSheetFiles(string file)
        {
            using (var fileStream = File.OpenRead(file))
            {
                using (var zipEntries = new ZipFile(fileStream))
                {
                    int counter = 0;

                    MemoryStream sheetStream = null;
                    MemoryStream stringStream = null;

                    foreach (ZipEntry zipEntry in zipEntries)
                    {
                        if (!zipEntry.IsFile)
                        {
                            continue; // Ignore directories
                        }

                        string entryFileName = zipEntry.Name;
                        if (entryFileName == FilesName.SheetName)
                        {
                            sheetStream = ZipEntryToStream(zipEntries, zipEntry);
                            sheetStream.Position = 0;
                            counter++;
                        }

                        if (entryFileName == FilesName.SharedStringsName)
                        {
                            stringStream = ZipEntryToStream(zipEntries, zipEntry);
                            stringStream.Position = 0;
                            counter++;
                        }

                        if (counter == 2)
                        {
                            // if both files read we can stop
                            break;
                        }
                    }

                    if (sheetStream is null || stringStream is null)
                    {
                        return null;
                    }

                    return new SheetEntry()
                    {
                        FileName = file,
                        StreamForSheetFile = sheetStream,
                        StreamForSharedStringsFile = stringStream
                    };
                }
            }
        }

        private MemoryStream ZipEntryToStream(ZipFile zf, ZipEntry zipEntry)
        {
            byte[] buffer = new byte[4096];     // 4K is optimum
            Stream zipStream = zf.GetInputStream(zipEntry);
            MemoryStream streamWriter = new MemoryStream();

            StreamUtils.Copy(zipStream, streamWriter, buffer);
            return streamWriter;

        }
    }
}