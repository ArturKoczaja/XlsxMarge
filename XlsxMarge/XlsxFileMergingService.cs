using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using ICSharpCode.SharpZipLib.Zip;

namespace XlsxMarge
{
    public class XlsxFileMergingService : IXlsxFileMergingService
    {
        private readonly SheetOperator _sheetOperator;
        private readonly DictionaryHelper _dictionaryHelper;
        private readonly FileOperator _fileOperator;
        private readonly XlsxFileExtractor _fileExtractor;

        public XlsxFileMergingService(
             SheetOperator sheetOperator,
             DictionaryHelper dictionaryHelper,
             FileOperator fileOperator,
             XlsxFileExtractor fileExtractor)
        {
            _sheetOperator = sheetOperator;
            _dictionaryHelper = dictionaryHelper;
            _fileOperator = fileOperator;
            _fileExtractor = fileExtractor;
        }

        public byte[] MergeFiles(List<byte[]> inputFilesBytes)
        {
            var files = _fileExtractor.UnzipXlsxFiles(inputFilesBytes);
            var allRows = new List<List<Cell>>();
            var addHeaders = true;

            // create dictionary
            foreach (var file in files)
            {
                var tmpDictionaryWithSharedStrings = _dictionaryHelper.CreateTmpSharedStringsDictionary(file);
                var rows = _sheetOperator.ReadRows(file);

                _sheetOperator.MergeRows(rows, ref addHeaders, allRows, tmpDictionaryWithSharedStrings);
            }

            var allSharedStrings = _dictionaryHelper.CreateSharedStringsDictionary(allRows);
            return PrepareOutputBytes(inputFilesBytes[0], allRows, allSharedStrings);
        }

        #region SheetMerging

        private byte[] PrepareOutputBytes(byte[] templateFileBytes, List<List<Cell>> allRows, Dictionary<string, long> allStrings)
        {
            using var fileStream = new MemoryStream();
            fileStream.Write(templateFileBytes, 0, templateFileBytes.Length);
            
            using var zipFile = new ZipFile(fileStream);

            using var sheetStream = _fileOperator.ReadFileToStream(zipFile, FileNames.SheetName);
            using var stringsStream = _fileOperator.ReadFileToStream(zipFile, FileNames.SharedStringsName);

            zipFile.BeginUpdate();

            _fileOperator.RemoveSheetAndStringsFiles(zipFile);

            using var outSheetStream = ReplaceSheetData(sheetStream, allRows, allStrings);
            using var outStringsStream = ReplaceStringsData(stringsStream, allStrings);

            var sheetDataSource = new XlsxMarge.CustomStaticDataSource();
            sheetDataSource.SetStream(outSheetStream);
            zipFile.Add(sheetDataSource, FileNames.SheetName);

            var stringsDataSource = new XlsxMarge.CustomStaticDataSource();
            stringsDataSource.SetStream(outStringsStream);
            zipFile.Add(stringsDataSource, FileNames.SharedStringsName);

            zipFile.CommitUpdate();

            return fileStream.ToArray();
        }

        private MemoryStream ReplaceSheetData(MemoryStream sheetStream, List<List<Cell>> allRows, Dictionary<string, long> allSharedStrings)
        {
            var xDocSheet = XDocument.Load(sheetStream);

            var sheetData = xDocSheet
                .Root
                .Descendants()
                .First(n => n.Name.LocalName == "sheetData");

            sheetData.RemoveAll();
            var xNamespace = xDocSheet.Root.Name.Namespace;
            var xNamespaceWithRow = xNamespace + "row";

            var columnNameArray = GetExcelColumnArray().ToArray();

            int rowCounter = 0;
            foreach (var row in allRows)
            {
                var rowElement = new XElement(xNamespaceWithRow, new XAttribute("r", rowCounter + 1), new XAttribute("spans", "1:15"), new XAttribute("ht", "12,75"));
                int columnCounter = 0;

                foreach (var col in row)
                {
                    var vElement = new XElement(xNamespace + "v", col.Translate ? allSharedStrings[col.Value] : col.Value);
                    var cElement = new XElement(xNamespace + "c", vElement,
                        new XAttribute("r", $"{columnNameArray[columnCounter]}{rowCounter + 1}"),
                        new XAttribute("s", col.Style),
                        new XAttribute("t", col.Translate ? "s" : "")
                    );

                    rowElement.Add(cElement);
                    columnCounter++;
                }

                sheetData.Add(rowElement);
                rowCounter++;
            }

            MemoryStream outputStream = new MemoryStream();
            xDocSheet.Save(outputStream);
            outputStream.Position = 0;
            return outputStream;
        }

        private static MemoryStream ReplaceStringsData(MemoryStream stringsStream, Dictionary<string, long> allSharedStrings)
        {
            var xDocStrings = XDocument.Load(stringsStream);
            xDocStrings.Root?.RemoveAll();
            var xNamespace = xDocStrings.Root.Name.Namespace;

            foreach (var str in allSharedStrings)
            {
                var tElement = new XElement(xNamespace + "t", str.Key);
                var siElement = new XElement(xNamespace + "si", tElement);

                xDocStrings.Root.Add(siElement);
            }

            MemoryStream outputStream = new MemoryStream();
            xDocStrings.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }

        static IEnumerable<string> GetExcelColumnArray()
        {
            string[] alphabet = { string.Empty, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            // TODO refactor 
            return from c1 in alphabet
                   from c2 in alphabet
                   from c3 in alphabet.Skip(1)                    // c3 is never empty
                   where c1 == string.Empty || c2 != string.Empty // only allow c2 to be empty if c1 is also empty
                   select c1 + c2 + c3;                 // c3 is never emptywhere c1 == string.Empty || c2 != string.Empty // only allow c2 to be empty if c1 is also emptyselect c1 + c2 + c3;
        }

        #endregion
    }
}