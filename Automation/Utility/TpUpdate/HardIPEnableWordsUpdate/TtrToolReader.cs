using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using OfficeOpenXml;

namespace Automation.Utility.TpUpdate.HardIPEnableWordsUpdate
{
    public class TtrToolReader
    {
        public Dictionary<string, EnableWordTable> ReadEnableWordTables(string path)
        {
            var tables = new Dictionary<string, EnableWordTable>();
            using (var ep = new ExcelPackage(new FileInfo(path)))
            {
                string regName = @"_\d+$";
                foreach (ExcelWorksheet ws in ep.Workbook.Worksheets)
                {
                    var enableWordTable = new EnableWordTable();
                    enableWordTable.ReadTable(ws);

                    if (!enableWordTable.IsValid)
                    {
                        continue;
                    }

                    string name = ws.Name;
                    if (Regex.IsMatch(ws.Name, regName, RegexOptions.IgnoreCase))
                    {
                        name = Regex.Replace(ws.Name, regName, "", RegexOptions.IgnoreCase);
                    }

                    if (!tables.ContainsKey(name) && enableWordTable.Rows.Count > 0)
                    {
                        tables.Add(ws.Name, enableWordTable);
                    }
                    else if (tables.TryGetValue(name, out EnableWordTable ewTable))
                    {
                        ewTable.EnableWords.AddRange(enableWordTable.EnableWords);
                        ewTable.EnableWords = ewTable.EnableWords.Distinct().ToList();
                        ewTable.Rows.AddRange(enableWordTable.Rows);
                    }
                }
                var enables = tables.Values.SelectMany(p => p.EnableWords).Distinct().ToList();
                tables.Values.ToList().ForEach(p => p.EnableWords = enables);
            }

            return tables;
        }
    }
}
