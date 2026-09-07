using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTable;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz
{
    public class BinTableSheetGenerator
    {
        protected readonly HardIpInputData HardIpInputData;

        private string BinTableSheetName
        {
            get
            {
                string binTableName = "Bin_Table_HardIP";
                switch (HardIpInputData.HardIpParaData.Block)
                {
                    case EnumBlock.HardIp:
                        binTableName = LocalSpecs.Options.Device == EnumDevice.RF ? "Bin_Table_ARF" : "Bin_Table_HardIP";
                        break;
                    case EnumBlock.Ids:
                        binTableName = "Bin_Table_HardIP";
                        break;
                    case EnumBlock.Rtos:
                        binTableName = "Bin_Table_Rtos";
                        break;
                }
                if (LocalSpecs.IsBenchLog)
                {
                    binTableName = "Bin_Table_ARF";
                }
                return binTableName;
            }
        }

        public BinTableSheetGenerator(HardIpInputData hardIpInputData)
        {
            HardIpInputData = hardIpInputData;
        }

        public BinTableSheet GenBinTable(Dictionary<string, HardIpSheet> planDic)
        {
            string path = LocalSpecs.IsBenchLog ? FolderStructure.TarDir : FolderStructure.DirHardIp;
            BinTableSheet existBintable = TestProgram.IgxlWorkBk.GetBinTblSheet(path, BinTableSheetName);
            BinTableSheet binTableSheet = existBintable.Rows.Any() ? existBintable : new BinTableSheet(BinTableSheetName);
            if (LocalSpecs.IsPatternValidate)
            {
                return binTableSheet;
            }

            var errorBinNums = new List<string>();
            var duplicateParameter = new List<string>();

            if (LocalSpecs.Options.Device == EnumDevice.AP)
            {
                var nonMnDict = new Dictionary<string, HardIpSheet>();
                var mnDict = new Dictionary<string, HardIpSheet>();

                foreach (KeyValuePair<string, HardIpSheet> plan in planDic)
                {
                    List<HardIpPattern> rows = plan.Value.Rows;
                    var nonMnRows = rows.Where(x => !x.Pattern.GetLastPayload().StartsWith("MN_", StringComparison.OrdinalIgnoreCase)).ToList();
                    var mnRows = rows.Where(x => x.Pattern.GetLastPayload().StartsWith("MN_", StringComparison.OrdinalIgnoreCase)).ToList();

                    nonMnDict[plan.Key] = new HardIpSheet { Rows = nonMnRows };
                    mnDict[plan.Key] = new HardIpSheet { Rows = mnRows };
                }
                GetBinTableByPlanDict(nonMnDict, binTableSheet, errorBinNums, duplicateParameter);
                GetBinTableByPlanDict(mnDict, binTableSheet, errorBinNums, duplicateParameter, mergedHln: true);
            }
            else
            {
                GetBinTableByPlanDict(planDic, binTableSheet, errorBinNums, duplicateParameter);
            }
            return binTableSheet;
        }
        private void GetBinTableByPlanDict(Dictionary<string, HardIpSheet> planDic, BinTableSheet binTableSheet, List<string> errorBinNums, List<string> duplicateParameter, bool mergedHln = false)
        {
            foreach (string sheetName in planDic.Keys)
            {
                BlockBinTableGeneratorBase blockBinGenerator;
                if (SearchInfo.IsHardipIdsSheet(sheetName))
                {
                    blockBinGenerator = new IdsBinTableGenerator(HardIpInputData, binTableSheet, sheetName, planDic[sheetName].Rows, duplicateParameter, errorBinNums);
                }
                else if (SearchInfo.IsHardipRtosSheet(sheetName))
                {
                    blockBinGenerator = new RtosBinTableGenerator(HardIpInputData, binTableSheet, sheetName, planDic[sheetName].Rows, duplicateParameter, errorBinNums);
                }
                else
                {
                    blockBinGenerator = new HardIpBinTableGenerator(HardIpInputData, binTableSheet, Regex.Replace(sheetName, "wireless_|lcd_", "", RegexOptions.IgnoreCase), planDic[sheetName].Rows, duplicateParameter, errorBinNums);
                }

                blockBinGenerator.GenerateBinTableRows(mergedHln);
            }
        }
    }
}
