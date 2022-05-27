using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using Common.Logging.Factory;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    public class XlsxMarge
    {
        private readonly string _sheetName = "xl/worksheets/sheet1.xml";
        private readonly string _sharedStringsName = "xl/sharedStrings.xml";

        //private readonly Dictionary<int, string> stringDictionary = new Dictionary<int, string>();

        private readonly string[] _inputFiles = new string[]
        {
            @"Res\test1.xlsx",
            @"Res\test2.xlsx"
        };

        private readonly string _outputFile = @"output.xlsx";

        static void Main(string[] args)
        {

            XlsxMarge mergeProgram = new XlsxMarge();
            mergeProgram.Run();

        }

        public class Cell
        {
            public Boolean Translate { get; set; }
            public string Value { get; set; }
        }

        private void Run()
        {

            var files = UnzipXlsxFiles(_inputFiles);

            List<List<Cell>> allRows = new List<List<Cell>>();


            Boolean addHeaders = true;

            // create dictionary
            foreach (var file in files)
            {
                var tmpDictionary = CreateTmpDictionary(file);

                // List<string> tmpRows = new List<string>();
                var rows = ReadRows(file);

                MergeRows(rows, ref addHeaders, allRows, tmpDictionary);
            }

            // WriteResultToConsole(allRows);

            var allStrings = CreateStringsDictionary(allRows);


            // WriteAllStringsToConsole(allStrings);

            PrepareOutputFile(_inputFiles[0], _outputFile);

        }

        private static void WriteAllStringsToConsole(Dictionary<string, long> allStrings)
        {
            foreach (var str in allStrings)
            {
                Console.WriteLine($"{str.Value}:\t\t\t{str.Key}");
            }

            Console.ReadKey();
        }

        private static Dictionary<string, long> CreateStringsDictionary(List<List<Cell>> allRows)
        {
            long cnt = 0;
            Dictionary<string, long> allStrings = new Dictionary<string, long>();
            foreach (var row in allRows)
            {
                foreach (var cell in row)
                {
                    if (cell.Translate)
                    {
                        if (!allStrings.ContainsKey(cell.Value))
                        {
                            allStrings.Add(cell.Value, cnt);
                            cnt++;
                        }
                    }
                }
            }

            return allStrings;
        }

        private bool MergeRows(IEnumerable<List<Cell>> rows, ref bool addHeaders, List<List<Cell>> allRows, Dictionary<int, string> tmpDictionary)
        {
            int cnt = 0;
            foreach (var row in rows)
            {
                if (addHeaders == false && cnt == 0)
                {
                }
                else
                {
                    allRows.Add(TranslateRow(row, tmpDictionary));
                    if (addHeaders == true)
                    {
                        addHeaders = false;
                    }
                }

                cnt++;
            }

            return addHeaders;
        }

        private static IEnumerable<List<Cell>> ReadRows(SheetEntry file)
        {
            XDocument xDocSheet;
            xDocSheet = XDocument.Load(file.SheetStream);

            var xxx = xDocSheet.Root.Descendants()
                .Where(n => n.Name.LocalName == "row");

            var rows = xxx.Select(row => row.Descendants()
                .Where(n => n.Name.LocalName == "c").Select(n => new Cell()
                {
                    Translate = (n as XElement).Attributes().FirstOrDefault(a => a.Name.LocalName == "t")?.Value == "s"
                            ? true
                            : false,
                    Value = (n as System.Xml.Linq.XElement).Value
                }
                )
                .ToList<Cell>());
            return rows;
        }

        private static Dictionary<int, string> CreateTmpDictionary(SheetEntry file)
        {
            XDocument xDocStrings;
            xDocStrings = XDocument.Load(file.StringStream);

            // to do convert to dictionary
            var strings = xDocStrings.Root.Descendants()
                .Where(n => n.Name.LocalName == "si")
                .Select((si, index) =>
                    new KeyValuePair<int, string>(index, (si.FirstNode as System.Xml.Linq.XElement).Value));

            var tmpDictionary = strings.ToDictionary(e => e.Key, e => e.Value);
            return tmpDictionary;
        }

        private static void WriteResultToConsole(List<List<Cell>> allRows)
        {
            foreach (var row in allRows)
            {
                foreach (var cell in row)
                {
                    Console.Write($"{cell.Value} ");
                }

                Console.WriteLine();
            }

            Console.ReadKey();
        }

        private List<Cell> TranslateRow(List<Cell> row, Dictionary<int, string> tmpDictionary)
        {
            List<Cell> result = new List<Cell>();
            foreach (var cell in row)
            {
                if (cell.Translate)
                {
                    bool _ = int.TryParse(cell.Value, out var index);
                    result.Add(new Cell()
                    {
                        Translate = cell.Translate,
                        Value = tmpDictionary[index]
                    });
                }
                else
                {
                    result.Add(new Cell()
                    {
                        Translate = cell.Translate,
                        Value = cell.Value
                    });
                }
            }

            return result;
        }


        public List<SheetEntry> UnzipXlsxFiles(string[] inputFiles, List<object> documentList = null)
        {
            List<SheetEntry> sheets = new List<SheetEntry>();
            foreach (var file in inputFiles)
            {
                sheets.Add(ExtractSheetFiles(file));

            }
            //foreach (var sheet in sheets)
            //{
            //    Console.WriteLine(sheet.FileName);
            //}
            return sheets;
        }

        private void PrepareOutputFile(string filePath, string outputPath)
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            File.Copy(filePath, outputPath);

            using FileStream fs = File.OpenRead(outputPath);
            using var zf = new ZipFile(fs);

            MemoryStream sheetStream = null;
            MemoryStream stringsStream = null;

            ReadFileToByteArrays(zf, ref sheetStream, ref stringsStream);

            RemoveSheetAndStringsFiles(zf);

            var outSheetStream = RemoveSheetData(sheetStream);
            var outStringsStream = RemoveStringsData(stringsStream);

           
            zf.BeginUpdate();

            var sheetDataSource = new CustomStaticDataSource();
            sheetDataSource.SetStream(outSheetStream);
            zf.Add(sheetDataSource, _sheetName);

            var stringsDataSource = new CustomStaticDataSource();
            stringsDataSource.SetStream(outStringsStream);
            zf.Add(stringsDataSource, _sharedStringsName);

            zf.CommitUpdate();
            
            //Console.Write(sheetStream);
            //Console.Write(stringsStream);

        }

        public class CustomStaticDataSource : IStaticDataSource
        {
            private Stream _stream;
            // Implement method from IStaticDataSource
            public Stream GetSource()
            {
                return _stream;
            }

            // Call this to provide the memorystream
            public void SetStream(Stream inputStream)
            {
                _stream = inputStream;
                _stream.Position = 0;
            }
        }

        private static MemoryStream RemoveSheetData(MemoryStream sheetStream)
        {
            XDocument xDocSheet;
            xDocSheet = XDocument.Load(sheetStream);

            var sheetData = xDocSheet.Root.Descendants()
                .First(n => n.Name.LocalName == "sheetData");

            sheetData.RemoveAll();

            MemoryStream outputStream = new MemoryStream();
            xDocSheet.Save(outputStream);
            // Rewind the stream ready to read from it elsewhere
            outputStream.Position = 0;
            return outputStream;
        }

        private static MemoryStream RemoveStringsData(MemoryStream stringsStream)
        {
            XDocument xDocStrings;
            xDocStrings = XDocument.Load(stringsStream);

            //var sheetData = xDocSheet.Root.Descendants()
            //    .First(n => n.Name.LocalName == "sst");

            xDocStrings.Root?.RemoveAll();

            MemoryStream outputStream = new MemoryStream();
            xDocStrings.Save(outputStream);
            // Rewind the stream ready to read from it elsewhere
            outputStream.Position = 0;
            return outputStream;
        }

        private void ReadFileToByteArrays(ZipFile zf, ref MemoryStream sheetStream, ref MemoryStream stringsStream)
        {
            foreach (ZipEntry zipEntry in zf)
            {
                if (!zipEntry.IsFile)
                {
                    continue; // Ignore directories
                }

                String entryFileName = zipEntry.Name;
                if (entryFileName == _sheetName)
                {
                    sheetStream = ZipEntryToStream(zf, zipEntry);
                    sheetStream.Position = 0;
                    //byte[] sheetData = stringsStream.ToArray();
                }

                if (entryFileName == _sharedStringsName)
                {
                    stringsStream = ZipEntryToStream(zf, zipEntry);
                    stringsStream.Position = 0;
                   // byte[] sheetData = stringsStream.ToArray();
                }
            }
        }

        private void RemoveSheetAndStringsFiles(ZipFile zf)
        {
            zf.BeginUpdate();

            foreach (ZipEntry zipEntry in zf)
            {
                if (!zipEntry.IsFile)
                {
                    continue; // Ignore directories
                }

                String entryFileName = zipEntry.Name;
                if (entryFileName == _sheetName)
                {
                    zf.Delete(zipEntry);
                }

                if (entryFileName == _sharedStringsName)
                {
                    zf.Delete(zipEntry);
                }
            }

            zf.CommitUpdate();
        }

        private SheetEntry ExtractSheetFiles(string file)
        {
            using (FileStream fs = File.OpenRead(file))
            {
                using (var zf = new ZipFile(fs))
                {
                    int cnt = 0;
                    MemoryStream sheetStream = null;
                    MemoryStream stringStream = null;

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
                            sheetStream = ZipEntryToStream(zf, zipEntry);
                            sheetStream.Position = 0;
                            cnt++;
                        }

                        if (entryFileName == _sharedStringsName)
                        {
                            stringStream = ZipEntryToStream(zf, zipEntry);
                            stringStream.Position = 0;
                            cnt++;
                        }

                        if (cnt == 2)
                        {
                            // if both files read we can stop
                            break;
                        }
                    }

                    if (sheetStream == null || stringStream == null) return null;
                    return new SheetEntry()
                    {
                        FileName = file,
                        SheetStream = sheetStream,
                        StringStream = stringStream
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

    public class SheetEntry
    {
        public string FileName { get; set; }
        public MemoryStream SheetStream { get; set; }
        public MemoryStream StringStream { get; set; }
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
