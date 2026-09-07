using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.Extension;

namespace TestPlanLib.Harvest
{
    public partial class CreateHarvestTruthTable(HarvestingTruthTableSheet harvestingTruthTableSheet, HarvestingFlagInitSheet harvestingFlagInitSheet, HarvestingFuseWriteSheet harvestingFuseWriteSheet, List<string> jobs)
    {
        [GeneratedRegex(@".*\[(?<value>.*)\]", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        [GeneratedRegex(@"^OR\((?<str>(\w+\.\w+(\,\w+\.\w+)+))\)", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex1();
        [GeneratedRegex(@"^Concatenate\((?<str>(\w+\.\w+(\+\w+\.\w+)+))\)", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex2();
        private static readonly Regex _regex = MyRegex();
        private readonly Regex _regexOr = MyRegex1();
        private readonly Regex _regexConcat = MyRegex2();

        private readonly HarvestingFlagInitSheet _flagInitSheet = harvestingFlagInitSheet;
        private readonly HarvestingFuseWriteSheet _fuseWriteSheet = harvestingFuseWriteSheet;
        private readonly HarvestingTruthTableSheet _baseSheet = harvestingTruthTableSheet;
        private readonly List<string> _jobs = [.. jobs.Select(x => x.ToUpper())];

        public List<HarvestingTruthTableSheet> Create()
        {
            var harvestingTruthTables = new List<HarvestingTruthTableSheet>();
            foreach (string job in _jobs)
            {
                HarvestingTruthTableSheet table = _baseSheet.Copy();
                table.Job = job;
                AddReadFuse(table, job);
                AddWriteFuse(table, job);
                harvestingTruthTables.Add(table);
            }
            return harvestingTruthTables;
        }

        private void AddReadFuse(HarvestingTruthTableSheet harvestingTruthTableSheet, string job)
        {
            if (_flagInitSheet == null)
            {
                return;
            }

            HarvestingFuseRead? flagInitTable = _flagInitSheet.EfuseMappingTable.Find(x => x.Job.EqualsIgnoreCase(job));
            Dictionary<string, string> blockDictionary = _flagInitSheet.BankDictionary;
            if (flagInitTable == null)
            {
                return;
            }

            foreach (KeyValuePair<string, string> item in flagInitTable.FuseMapping)
            {
                string flag = item.Key;
                if (item.Value == "0" || string.IsNullOrEmpty(flag) || string.IsNullOrEmpty(item.Value))
                {
                    continue;
                }

                if (_regexOr.IsMatch(item.Value))
                {
                    foreach (HarvestingTruthTableRow row in harvestingTruthTableSheet.Rows)
                    {
                        row.ReadFuse.Add(AddMultiFuseRead(flag, item.Value));
                    }
                    continue;
                }
                string[] fuseItem = item.Value.Split('.');
                string blockName = "";
                string fuseName = "";
                if (fuseItem.Length > 1)
                {
                    blockName = fuseItem.FirstOrDefault()!;
                    fuseName = fuseItem.Last();
                }
                else
                {
                    fuseName = fuseItem.FirstOrDefault()!;
                    if (!blockDictionary.TryGetValue(fuseName, out string? value))
                    {
                        continue;
                    }
                    else
                    {
                        blockName = value;
                    }
                }
                if (flag.StartsWith("F_BinFuse"))
                {
                    string bitValue = _regex.Match(flag).Groups["value"].ToString();
                    harvestingTruthTableSheet.ReadBinFuseBit = int.Parse(bitValue.Split(':').First()) - int.Parse(bitValue.Split(':').Last()) + 1;
                }
                foreach (HarvestingTruthTableRow row in harvestingTruthTableSheet.Rows)
                {
                    row.ReadFuse.Add(new Fusing { Field = blockName.ToLower().Replace("bank_", ""), Address = string.Format(fuseName + "=>" + flag) });
                }
                harvestingTruthTableSheet.IndexReadFuse = 1;
            }
        }

        private Fusing AddMultiFuseRead(string flag, string fuseItem)
        {
            string fuseValue = _regexOr.Match(fuseItem).Groups["str"].ToString();
            IEnumerable<string> fieldList = fuseValue.Split(',').ToList().Select(x => x.Split('.').Last());
            IEnumerable<string> bankList = fuseValue.Split(',').ToList().Select(x => x.Split('.').First().ToLower().Replace("bank_", ""));
            string field = $"OR({string.Join(",", fieldList)})";
            string bank = $"OR({string.Join(",", bankList)})";
            return new Fusing { Field = bank, Address = $"{field}=>{flag}" };
        }

        private void AddWriteFuse(HarvestingTruthTableSheet harvestingTruthTableSheet, string job)
        {
            if (_fuseWriteSheet == null)
            {
                return;
            }

            List<HarvestingFuseWriteRow> fuseWriteSheet = _fuseWriteSheet.Rows.FindAll(x => x.Job.EqualsIgnoreCase(job));
            if (fuseWriteSheet.Count == 0)
            {
                return;
            }

            foreach (HarvestingFuseWriteRow item in fuseWriteSheet)
            {
                string[] fuseItem = item.FuseName.Split('.');
                string blockName = item.BlockName;
                string fuseName = "";
                if (fuseItem.Length > 1)
                {
                    blockName = fuseItem.FirstOrDefault()!;
                    fuseName = fuseItem.Last();
                }
                else
                {
                    fuseName = fuseItem.FirstOrDefault()!;
                }
                string value = item.Value;
                if (_regexConcat.IsMatch(value))
                {
                    string oriFlagValue = _regexConcat.Match(value).Groups["str"].ToString();
                    value = $"Concatenate({string.Join(" +", oriFlagValue.Split('+').ToList().Select(x => $"{x.Split('.').First().ToLower().Replace("bank_", "")}.{x.Split('.').Last()}"))})";
                }
                if (item.Value.StartsWith("F_BinFuse"))
                {
                    string bitValue = _regex.Match(value).Groups["value"].ToString();
                    harvestingTruthTableSheet.SetBinFuseBit = int.Parse(bitValue.Split(':').First()) - int.Parse(bitValue.Split(':').Last()) + 1;
                }
                foreach (HarvestingTruthTableRow row in harvestingTruthTableSheet.Rows)
                {
                    row.Fusings.Add(new Fusing { Field = blockName.ToLower().Replace("bank_", ""), Address = fuseName, Value = value });
                }
            }
        }
    }
}
