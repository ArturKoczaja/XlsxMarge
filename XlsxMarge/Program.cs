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
            var xlsxFileExtractor = new XlsxFileExtractor();
            var files = xlsxFileExtractor.UnzipXlsxFiles(_inputFiles);

            var xlsxFileMergingService = new XlsxFileMergingService(new SheetOperator(), new DictionaryHelper(), new FileOperator());
            xlsxFileMergingService.MergeFiles(files, _inputFiles[0], _outputFile);
        }
    }
}
