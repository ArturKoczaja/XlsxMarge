namespace XlsxMarge
{
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