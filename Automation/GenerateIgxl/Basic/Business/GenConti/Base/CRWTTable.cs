using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base
{
    public class CrwtRow
    {
        public string PinGroup { get; }
        public string IRange { get; }
        public string WaitTime { get; }

        public CrwtRow(string pinGroup, string iRange, string waitTime)
        {
            PinGroup = pinGroup;
            IRange = iRange;
            WaitTime = waitTime;
        }
    }

    public class CrwtTable
    {
        private string _defaultIRange = "0.02";
        private string _defaultWaitUf = "0.00054";
        private string _defaultWaitUfp = "0.0002";

        public static readonly string[] Header =
        {
            "Pin",
            "CurrentRange",
            "WaitTime"
        };

        public List<CrwtRow> Rows { get; } = new List<CrwtRow>();

        public CrwtTable(List<ContiBase> contiBases, string platform)
        {
            string waitTime = platform == "UF" ? _defaultWaitUf : _defaultWaitUfp;

            var pinGroups = contiBases
                .Where(x => x is ContiPowerShort)
                .Select(x => x.DcTestContiRow?.PinGroup ?? string.Empty)
                .Where(pg => !string.IsNullOrWhiteSpace(pg))
                .SelectMany(pg => pg.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                .Select(pg => pg.Trim())
                .Where(pg => pg.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(pg => pg, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (string pg in pinGroups)
            {
                Rows.Add(new CrwtRow(pg, _defaultIRange, waitTime));
            }
        }

        public void ExportToTxt(string path)
        {
            Directory.CreateDirectory(path);

            var sb = new StringBuilder();
            sb.AppendLine(string.Join("\t", Header));
            foreach (CrwtRow row in Rows)
            {
                sb.AppendLine($"{row.PinGroup}\t{row.IRange}\t{row.WaitTime}");
            }

            const string outputFile = "TuneCRWTResult.txt";
            File.WriteAllText(Path.Combine(path, outputFile), sb.ToString(), Encoding.Default);
        }

    }

}
