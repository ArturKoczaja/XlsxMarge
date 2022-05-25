﻿using System;
using System.Collections.Generic;
using System.IO;

namespace XlsxMarge
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string[] inputFiles = new string[]
            {
                @"C:\tfs\trening\excelFile\1RX3015_20220512_Testberekening - XLSX_1.xlsx",
                @"C:\tfs\trening\excelFile\1RX3015_20220512_Testberekening - XLSX_2.xlsx"
            };

            string[] outputFiles = new string[]
            {
                @"C:\tfs\trening\excelFile\output.xlsx"
            };

            // Load xlsx files as byte[] and add them to documentItemList. Do it in separated static method.
            var documentItemList = new List<DocumentItem>();

            var xlsXlsArchiveService = new XlsArchiveService();
            var xlsxFileMergingService = new XlsxFileMergingService(xlsXlsArchiveService);
            var result = xlsxFileMergingService.MergeFiles(documentItemList);
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
        // in documentStore the same filename can be represented by multiple files
        public string Filename { get; set; }

        // Collection of multiple data
        // byte[] is single file
        public IEnumerable<byte[]> Data { get; set; }
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

