using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace XlsxMarge
{
    public partial class XlsxMarge
    {
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

        private void Run()
        {
            var bytes = new List<byte[]>();
            foreach (var path in _inputFiles)
            {
                var byteEntry = File.ReadAllBytes(path);
                bytes.Add(byteEntry);
            }
            
            var xlsxFileMergingService = new XlsxFileMergingService(new SheetOperator(), new DictionaryHelper(), new FileOperator(), new XlsxFileExtractor());
            var outBytes = xlsxFileMergingService.MergeFiles(bytes);

            if (File.Exists(_outputFile))
            {
                File.Delete(_outputFile);
            }

            using (var outputFileStream = File.Open(_outputFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                outputFileStream.Write(outBytes, 0, outBytes.Length);
            }

            Console.WriteLine("Finito.");
            Console.ReadKey();

        }
    }
}
