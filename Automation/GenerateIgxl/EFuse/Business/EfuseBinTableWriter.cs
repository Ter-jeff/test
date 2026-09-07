using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Static;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.EFuse.Business
{
    internal class EfuseBinTableWriter
    {
        public void GenBinTable(List<string> existEfuseFlags)
        {
            BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet(FolderStructure.DirBinTable);
            List<BinTableRow> binTableRows = GenerateBinTableRows(existEfuseFlags);
            binTable.AddRows(binTableRows);

            // MainInitEnableWd
            var flags = binTableRows.SelectMany(x => x.ItemList.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, EFuseConst.Efuse, FolderStructure.DirMain);
        }

        private List<BinTableRow> GenerateBinTableRows(List<string> existEfuseFlags)
        {
            var binTableRows = new List<BinTableRow>();
            BinNumberSingletonLegacy.Instance().SetStartBinNumber(90);
            var blankOrder = new List<string>
            {
                BankType.Ecid,
                BankType.Cfg,
                BankType.UdrP,
                BankType.UdrP0,
                BankType.UdrP1,
                BankType.UdrE,
                BankType.UdrE0,
                BankType.UdrE1,
                BankType.Mon,
                BankType.Unknow
            };
            IOrderedEnumerable<IGrouping<string, string>> groups = existEfuseFlags.GroupBy(x => EFuseConst.GetBankName(x)).OrderBy(y => blankOrder.IndexOf(y.Key));
            foreach (IGrouping<string, string> group in groups)
            {
                var newGroup = new List<string>();
                if (group.Key.Equals(BankType.Cfg))
                {
                    newGroup.AddRange(group.Where(x => x.EndsWith(EFuseConst.Early)));
                    newGroup.AddRange(group.Where(x => !x.EndsWith(EFuseConst.Early)));
                }
                else if (group.Key.Equals(BankType.Ecid))
                {
                    newGroup.AddRange(group.Where(x => !x.EndsWith(EFuseConst.Deid) && !x.EndsWith(EFuseConst.NonDeid)));
                    newGroup.AddRange(group.Where(x => x.EndsWith(EFuseConst.NonDeid)));
                }

                foreach (string item in newGroup.Any() ? (IEnumerable<string>)newGroup : group)
                {
                    var row = new BinTableRow();
                    string flagName = item;
                    string name = Regex.Replace(item, "efuse_", "", RegexOptions.IgnoreCase);
                    string binTableName = name.Replace("F_", "Bin_EFUSE_");

                    if (binTableName.ContainsIgnoreCase("BANKREAD") || binTableName.ToUpper().EndsWith("JUDGE_DRAM_TYPE"))
                    {
                        binTableName = binTableName.Replace("EFUSE", "EFUSE_RT");
                    }

                    if (item.Equals("F_" + EFuseConst.PseudoFuseWriteItem) && !(TestPlanStatic.EfuseInstanceSheets != null && TestPlanStatic.EfuseInstanceSheets.Any()))
                    {
                        binTableName = "Bin_" + EFuseConst.PseudoFuseWriteItem;
                    }
                    string[] binTableNameSeg = binTableName.Split('_');
                    string category1 = binTableNameSeg.Length > 2 ? binTableNameSeg[2] : "";
                    string category2 = binTableNameSeg.Length > 3 ? binTableNameSeg[3] : "";

                    row.Name = binTableName;
                    row.ItemList = flagName;
                    row.Items.Add("T");
                    row.Op = "AND";
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Efuse", category1, category2, row);
                    row.Sort = binNumInfo.SoftBin.ToString("G15");
                    row.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    row.Result = binNumInfo.BinNumInfo.Status;
                    binTableRows.Add(row);
                }
            }
            return binTableRows;
        }

        public BinTableRow GenBinTableRow(string name)
        {
            var row = new BinTableRow { Name = name, ItemList = name.Replace("Bin_EFUSE_", "F_") };
            row.Items.Add("T");
            row.Op = "AND";
            BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Efuse", "", "", row);
            row.Sort = binNumInfo.SoftBin.ToString("G15");
            row.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
            row.Result = binNumInfo.BinNumInfo.Status;
            return row;
        }
    }
}
