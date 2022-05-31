using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace XlsxMarge
{
    public class DictionaryHelper
    {
        public Dictionary<int, string> CreateTmpSharedStringsDictionary(SheetEntry file)
        {
            var stringStream = new MemoryStream(file.stringsBytes);
            var xDocSharedStrings = XDocument.Load(stringStream);
            IEnumerable<KeyValuePair<int, string>> strings = xDocSharedStrings
                .Root
                .Descendants()
                .Where(n => n.Name.LocalName == "si")
                .Select((si, index) => new KeyValuePair<int, string>(index, (si.FirstNode as XElement).Value));

            var tmpDictionary = strings.ToDictionary(e => e.Key, e => e.Value);
            return tmpDictionary;
        }

        public Dictionary<string, long> CreateSharedStringsDictionary(List<List<Cell>> allRows)
        {
            long counter = 0;
            var allSharedStrings = new Dictionary<string, long>();

            foreach (var row in allRows)
            {
                foreach (var cell in row)
                {
                    if (cell.Translate)
                    {
                        if (!allSharedStrings.ContainsKey(cell.Value))
                        {
                            allSharedStrings.Add(cell.Value, counter);
                            counter++;
                        }
                    }
                }
            }

            return allSharedStrings;
        }
    }
}