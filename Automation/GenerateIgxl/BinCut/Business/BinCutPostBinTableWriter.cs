using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutPostBinTableWriter : BinCutBinTableWriter
    {
        public void GetOutsideBinTable(Dictionary<string, List<BinCutSourceItem>> sourceRowDic, BinCutInputData binCutInputManager)
        {
            var binTableBinCutSheet = new BinTableSheet("Bin_Table_BinCut_Outside");
            var list = new List<string>();
            foreach (BinCutFlowTables flowSheet in binCutInputManager.BinCutPostFlowSheets)
            {
                list.AddRange(flowSheet.GetPerfromanceModeInFlow(EnumBinCutTableType.Post));
            }

            list = list.Distinct().ToList();
            binTableBinCutSheet.Rows.AddRange(GenOutsideBinTableRows(list));
            TestProgram.IgxlWorkBk.AddBinTblSheet(FolderStructure.DirBinCut, binTableBinCutSheet);
            var flags = binTableBinCutSheet.Rows.SelectMany(x => x.ItemList.Split(',')).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            TestProgram.IgxlWorkBk.GenAllFailFlagToMainInitEnableWd(flags, "BinCut", FolderStructure.DirMain);
        }

        private IEnumerable<BinTableRow> GenOutsideBinTableRows(List<string> modes)
        {
            var binTableRows = new List<BinTableRow>();

            var levelList = new List<EnumColumnName> { EnumColumnName.TD, EnumColumnName.Mbist, EnumColumnName.ILB, EnumColumnName.ELB };
            if (string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder))
            {
                levelList.Add(EnumColumnName.RTOS);
            }

            foreach (string mode in modes)
            {
                foreach (EnumColumnName level in levelList)
                {
                    EnumColumnName domain = level;
                    var row = new BinTableRow { Name = "Bin_" + domain + "_" + mode + "_outsidebincut_BV" };
                    string binName = "F_" + domain + "_" + mode + "_outsidebincut_BV";
                    row.ItemList = binName;
                    row.Items.Add("T");
                    row.Op = "AND";
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", "PBC", "", row);
                    row.Result = binNumInfo.BinNumInfo.Status;
                    row.Bin = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                    row.Sort = binNumInfo.SoftBin.ToString("G15");
                    binTableRows.Add(row);
                }
            }
            return binTableRows;
        }
    }
}
