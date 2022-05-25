using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using System.Runtime.CompilerServices;
using System.Xml;
using Common.Logging;
using Excel = Microsoft.Office.Interop.Excel;
using Common.Logging.Factory;

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

            Excel.Application app = new Excel.Application();
            app.Workbooks.Add(inputFiles);
            app.Workbooks.Add(outputFiles);
            ZipDirectory(@"C:\tfs\trening\excelFile\1RX3015_20220512_Testberekening - XLSX_result.xlsx", "output.xlsx");
            UnzipFile("1RX3015_20220512_Testberekening - XLSX_1.xlsx", "TemporaryDirectory");
            
        }

    

        public static void UnzipFile(string zipFileName, string targetDirectory)
        {
            new FastZip().ExtractZip(zipFileName, targetDirectory, null);
        }

        public static void ZipDirectory(string sourceDirectory, string zipFileName)
        {
            new FastZip().CreateZip(zipFileName, sourceDirectory, true, null);
        }

        public static IList<string> ReadStringTable(Stream inputFile)
        {
            var stringTable = new List<string>();

            using (var reader = XmlReader.Create(inputFile))
                for (reader.MoveToContent(); reader.Read();)
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "t")
                        stringTable.Add(reader.ReadElementString());

            return stringTable;
        }

        public static void WriteStringTable(Stream outputFile, IList<string> stringTable)
        {
            using (var writer = XmlWriter.Create(outputFile))
            {
                writer.WriteStartDocument(true);

                writer.WriteStartElement("sst", "");
                writer.WriteAttributeString("count", stringTable.Count.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("uniqueCount", stringTable.Count.ToString(CultureInfo.InvariantCulture));

                foreach (var str in stringTable)
                {
                    writer.WriteStartElement("si");
                    writer.WriteElementString("t", str);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }
        }

        public static void ReadWorksheet(Stream inputFile, IList<string> stringTable, Excel.DataTable data)
        {
            using (var reader = XmlReader.Create(inputFile))
            {
                DataRow row = null;
                int columnIndex = 0;
                string type;
                int value;

                for (reader.MoveToContent(); reader.Read();)
                    if (reader.NodeType == XmlNodeType.Element)
                        switch (reader.Name)
                        {
                            case "c":
                                type = reader.GetAttribute("t");
                                reader.Read();
                                value = int.Parse(reader.ReadElementString(), CultureInfo.InvariantCulture);

                                if (type == "s")
                                    row[columnIndex] = stringTable[value];
                                else
                                    row[columnIndex] = value;

                                columnIndex++;

                                break;
                        }
            }
        }

        public static IList<string> CreateStringTables(DataTable data, out IDictionary<string, int> lookupTable)
        {
            var stringTable = new List<string>();
            lookupTable = new Dictionary<string, int>();

            foreach (DataRow row in data.Rows)
            foreach (DataColumn column in data.Columns)
                if (column.DataType == typeof(string))
                {
                    var value = (string) row[column];

                    if (!lookupTable.ContainsKey(value))
                    {
                        lookupTable.Add(value, stringTable.Count);
                        stringTable.Add(value);
                    }
                }

            return stringTable;

        }
    }

}

