using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.TempMon.Data;

namespace Automation.GenerateIgxl.PostAction.TempMon.Writer
{
    public class TempMonConditionSheet
    {
        public const string SheetName = "TempMonExecuteSetting";
        public IEnumerable<string> Headers { get; } = new List<string>()
        {
            "Mode",
            "Condition",
            "Type",
            "Test Item"
        };

        public HashSet<TempMonData> Datas { get; }

        public TempMonConditionSheet(HashSet<TempMonData> datas)
        {
            Datas = datas;
        }

        public void Write(string path)
        {
            Directory.CreateDirectory(path);
            List<string> content = new List<string>();
            content.Add(GetHeadersLine());
            content.AddRange(GetDatasLines());
            File.WriteAllLines(Path.Combine(path, $"{SheetName}.txt"), content);
        }

        private string GetHeadersLine()
        {
            return string.Join("\t", Headers);
        }

        private List<string> GetDatasLines()
        {
            return Datas.Select(x => $"{x.Mode}\t{x.Condition}\t{x.Type}\t{x.Item}").ToList();
        }
    }
}
