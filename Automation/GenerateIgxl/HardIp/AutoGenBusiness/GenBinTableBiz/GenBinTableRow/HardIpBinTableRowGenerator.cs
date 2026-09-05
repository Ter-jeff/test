using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow
{
    public class HardIpBinTableRowGenerator : BinTableRowGeneratorBase
    {
        private const string Opcode = "OR";
        public HardIpBinTableRowGenerator(string sheetName, List<string> errorBinNums)
            : base(sheetName, errorBinNums)
        {
        }

        protected override BinTableRow GenBinTableRowForPattern(string voltage = "")
        {

            var binTableRow = new BinTableRow
            {
                Name = CreateHardIpName(),
                ItemList = CreateHardIpItemList(voltage),
                Op = CreateHardIpOpCode(),
                Items = CreateHardIpItems(voltage),
                Sort = CreateSortBin(),
                Bin = CreateHardBin(),
                Result = CreateResult()
            };
            //binTableRow.ExtraBinDictionary = CreateSortBinExtraBinDic();
            //CheckErrorBinNum();
            return binTableRow;
        }

        protected override void SetPattern(HardIpPattern pattern)
        {
            SetBasicInfoByPattern(pattern);
        }

        internal virtual string CreateHardIpItemList(string voltage = "")
        {
            List<string> flagVoltageList = new List<string>();
            string failFlagUserDefineN = "";
            string failFlagUserDefineH = "";
            string failFlagUserDefineL = "";
            if (IsHipEfuseRead)
            {
                return CommonGenerator.GenHipEfuseReadTestFailAction();
            }

            if (!string.IsNullOrEmpty(Pattern.Failflag))
            {
                flagVoltageList = Pattern.Failflag.Split(new[] { ';', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }

            string patternName = Pattern.Pattern.GetPatternName();
            if (flagVoltageList.Exists(x => x.StartsWith(HardIpConstData.LabelNv)))
            {
                failFlagUserDefineN = flagVoltageList.Find(x => x.StartsWith(HardIpConstData.LabelNv)).Split('@').Last();
            }
            if (flagVoltageList.Exists(x => x.StartsWith(HardIpConstData.LabelHv)))
            {
                failFlagUserDefineH = flagVoltageList.Find(x => x.StartsWith(HardIpConstData.LabelHv)).Split('@').Last();
            }
            if (flagVoltageList.Exists(x => x.StartsWith(HardIpConstData.LabelLv)))
            {
                failFlagUserDefineL = flagVoltageList.Find(x => x.StartsWith(HardIpConstData.LabelLv)).Split('@').Last();
            }

            string failFlagDefaultN = LocalSpecs.Options.Device == EnumDevice.RF ? CommonGenerator.GenWirelessFlowFailFlag(BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), HardIpConstData.LabelNv) :
                CommonGenerator.GenHardIpFlowFailFlag(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), patternName.Replace(PatternClass.MultipleConst, ""), TimingAc, InstNameSubStr, HardIpConstData.LabelNv, NoPattern);
            string failFlagDefaultH = LocalSpecs.Options.Device == EnumDevice.RF ? CommonGenerator.GenWirelessFlowFailFlag(BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), HardIpConstData.LabelHv) :
                CommonGenerator.GenHardIpFlowFailFlag(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), patternName.Replace(PatternClass.MultipleConst, ""), TimingAc, InstNameSubStr, HardIpConstData.LabelHv, NoPattern);
            string failFlagDefaultL = LocalSpecs.Options.Device == EnumDevice.RF ? CommonGenerator.GenWirelessFlowFailFlag(BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), HardIpConstData.LabelLv) :
                CommonGenerator.GenHardIpFlowFailFlag(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), patternName.Replace(PatternClass.MultipleConst, ""), TimingAc, InstNameSubStr, HardIpConstData.LabelLv, NoPattern);

            string failFlagN = string.IsNullOrEmpty(failFlagUserDefineN) ? failFlagDefaultN : failFlagUserDefineN;
            string failFlagH = string.IsNullOrEmpty(failFlagUserDefineH) ? failFlagDefaultH : failFlagUserDefineH;
            string failFlagL = string.IsNullOrEmpty(failFlagUserDefineL) ? failFlagDefaultL : failFlagUserDefineL;

            var flagList = new List<string>();
            if (voltage == "")
            {
                flagList = new List<string> { failFlagN, failFlagH, failFlagL };
            }
            else
            {
                if (voltage == HardIpConstData.LabelNv)
                {
                    flagList.Add(failFlagN);
                }

                if (voltage == HardIpConstData.LabelHv)
                {
                    flagList.Add(failFlagH);
                }

                if (voltage == HardIpConstData.LabelLv)
                {
                    flagList.Add(failFlagL);
                }
            }
            return string.Join(",", flagList);
        }

        internal string CreateHardIpName()
        {
            if (IsHipEfuseRead)
            {
                return CommonGenerator.GenHardIpEfuseReadBinParameter(SheetName);
            }

            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""));
        }

        private string CreateHardIpOpCode()
        {
            return Opcode;
        }

        internal List<string> CreateHardIpItems(string voltage = "")
        {
            switch (voltage)
            {
                case HardIpConstData.LabelHv:
                    return new List<string> { "", "T", "" };
                case HardIpConstData.LabelLv:
                    return new List<string> { "", "", "T" };
                case HardIpConstData.LabelNv:
                    return new List<string> { "T", "", "" };
                default:
                    return new List<string> { "T", "T", "T" };
            }
        }
    }
}
