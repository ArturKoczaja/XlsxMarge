using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace XlsxMarge
{
    public class SheetOperator
    {
        public IEnumerable<List<Cell>> ReadRows(SheetEntry file)
        {

            string xmlText = Encoding.ASCII.GetString(file.sheetBytes);
            var xDocSheet = XDocument.Parse(xmlText);
            var xxx = xDocSheet
                .Root
                .Descendants()
                .Where(n => n.Name.LocalName == "row");

            var rows = xxx
                .Select(row => row.Descendants()
                    .Where(n => n.Name.LocalName == "c")
                    .Select(n => new Cell()
                    {
                        Translate = n.Attributes().FirstOrDefault(a => a.Name.LocalName == "t")?.Value == "s",
                        Value = n.Value
                    })
                    .ToList());

            return rows;
        }

        public void MergeRows(
            IEnumerable<List<Cell>> rows,
            ref bool addHeaders,
            List<List<Cell>> allRows,
            Dictionary<int, string> tmpDictionary)
        {
            int counter = 0;

            foreach (var row in rows)
            {
                if (addHeaders is false && counter is 0)
                {
                }
                else
                {
                    allRows.Add(TranslateRow(row, tmpDictionary));
                    if (addHeaders)
                    {
                        addHeaders = false;
                    }
                }

                counter++;
            }
        }

        private List<Cell> TranslateRow(List<Cell> row, Dictionary<int, string> tmpDictionary)
        {
            var result = new List<Cell>();

            foreach (var cell in row)
            {
                if (cell.Translate)
                {
                    int.TryParse(cell.Value, out var index);
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
    }
}