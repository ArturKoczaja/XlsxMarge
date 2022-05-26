using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    internal class Program
    {
        private static readonly string _sheetName = "xl/worksheets/sheet1.xml";
        private static readonly string _sharedStringsName = "xl/sharedStrings.xml";

        static void Main(string[] args)
        {
            string[] inputFiles = new string[]
            {
                @"Res\test1.xlsx",
                @"Res\test2.xlsx"
            };

            string[] outputFiles = new string[]
            {
                @"output.xlsx"
            };

            // Load xlsx files as byte[] and add them to documentItemList. Do it in separated static method.
            //var documentItemList = new List<DocumentItem>();

            UnzipXlsxFiles(inputFiles);

            Console.ReadKey();
            //var xlsXlsArchiveService = new XlsArchiveService();
            //var xlsxFileMergingService = new XlsxFileMergingService(xlsXlsArchiveService);
            //var result = xlsxFileMergingService.MergeFiles(documentItemList);
        }


        public static void UnzipXlsxFiles(string[] inputFiles, List<object> documentList = null)
        {

            foreach (var file in inputFiles)
            {
                using (FileStream fs = File.OpenRead(file))
                {
                    using (var zf = new ZipFile(fs))
                    {
                        foreach (ZipEntry zipEntry in zf)
                        {
                            if (!zipEntry.IsFile)
                            {
                                continue; // Ignore directories
                            }
                           
                            String entryFileName = zipEntry.Name;
                            //Console.WriteLine(entryFileName);
                            if (entryFileName == _sheetName)
                            {
                                Console.WriteLine(entryFileName);
                            }
                            if (entryFileName == _sharedStringsName)
                            {
                                Console.WriteLine(entryFileName);

                            }
                        }
                    }
                }

               
            }
        }

        public static void LoadXlsxFiles(string[] inputFiles, List<DocumentItem> documentList)
        {
            foreach (var file in inputFiles)
            {
                var loadFile = File.ReadAllBytes(file);
                documentList.Add(new DocumentItem()
                {
                    Data = loadFile,
                    Filename = file
                });
            }
        }
    }

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

    public interface IXlsArchiveService
    {
        IEnumerable<XlsFileToMerge> UnzipFiles(IEnumerable<DocumentItem> documentItems);
        Stream ZipFilesAsStream(object obj);
    }

    // Will be used in document store to merge all files
    public class XlsxFileMergingService : IXlsxFileMergingService
    {
        private readonly IXlsArchiveService _xlsArchiveService;

        public XlsxFileMergingService(IXlsArchiveService xlsArchiveService)
        {
            _xlsArchiveService = xlsArchiveService;
        }

        public Stream MergeFiles(IEnumerable<DocumentItem> documentItems)
        {
            var xlsFilesToMerge = _xlsArchiveService.UnzipFiles(documentItems);

            //...
            // Logic for merging xlsx files 
            //...

            var mergedFilesAsStream = _xlsArchiveService.ZipFilesAsStream(xlsFilesToMerge);
            return mergedFilesAsStream;
        }
    }

    public interface IXlsxFileMergingService
    {
        Stream MergeFiles(IEnumerable<DocumentItem> documentItems);
    }


    public class DocumentItem
    {
        public string Filename { get; set; }
        public byte[] Data { get; set; }
    }

    public class XlsFileToMerge
    {
        public XlsFileToMerge(string name, byte[] sharedStringFile, byte[] sheetFile)
        {
            Name = name;
            SharedStringFile = sharedStringFile;
            SheetFile = sheetFile;
        }

        public string Name { get; }

        public byte[] SharedStringFile { get; }

        public byte[] SheetFile { get; }
    }
}
