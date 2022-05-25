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

            ZipDirectory(@"C:\tfs\trening\excelFile\1RX3015_20220512_Testberekening - XLSX_result.xlsx", "output.xlsx");
            UnzipFile("1RX3015_20220512_Testberekening - XLSX_1.xlsx", "TemporaryDirectory");
            
        }
}

