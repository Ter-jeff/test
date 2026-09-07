using System;
using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Static;
using Automation.Utility.HardIP;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow
{
    public class RtosBinTableRowGenerator : BinTableRowGeneratorBase
    {
        private const string RtosOpcode = "OR";
        private bool _idsNoFuse;
        private string _repeatStr;
        public RtosBinTableRowGenerator(string sheetName, List<string> errorBinNums)
            : base(sheetName, errorBinNums)
        {

        }

        protected override BinTableRow GenBinTableRowForPattern(string voltage = "")
        {
            var binTableRow = new BinTableRow
            {
                ItemList = CreateRtosItemList(),
                Name = CreateRtosName(),
                Items = CreateRtosItems(),
                Op = CreateRtosOpcode(),
                Sort = CreateSortBin(),
                Bin = CreateHardBin(),
                Result = CreateResult()
            };
            return binTableRow;
        }

        protected override void SetPattern(HardIpPattern pattern)
        {
            SetBasicInfoByPattern(pattern);
            _idsNoFuse = HardIpService.IdsNoFuse(pattern.MiscInfo);
            _repeatStr = HardIpService.GetRepeatMapping(pattern.MiscInfo);
        }

        private string CreateRtosItemList()
        {
            if (!string.IsNullOrEmpty(Pattern.FunctionName) && !Pattern.FunctionName.Equals(FuncNameConst.CSharpFuncNameIdsCurrent, StringComparison.OrdinalIgnoreCase))
            {
                return string.IsNullOrEmpty(Pattern.Failflag) ? $"{HardIpConstData.PrefixHardIpFailAction}_{BlockName}_{SubBlockName}{HardIpConstData.SuffixHardIpFailAction}" : Pattern.Failflag;
            }
            if (!string.IsNullOrEmpty(Pattern.FunctionName) && Pattern.FunctionName.Equals(FuncNameConst.CSharpFuncNameIdsCurrent, StringComparison.OrdinalIgnoreCase))
            {
                if (_idsNoFuse)
                {
                    return string.IsNullOrEmpty(Pattern.Failflag) ? $"{HardIpConstData.PrefixHardIpFailAction}_{BlockName}_{SubBlockName}{HardIpConstData.SuffixHardIpFailAction}" : Pattern.Failflag;
                }

                if (_repeatStr != "")
                {
                    var repeatList = new List<string>();
                    foreach (string repeatItem in _repeatStr.Split(','))
                    {
                        repeatList.Add("F_IDS_Current_Main" + "_" + repeatItem.Replace(".", "p"));
                    }
                    return string.Join(",", repeatList);
                }
                return "F_IDS_Current_Main_1x";
            }
            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName);
        }

        private string CreateRtosName()
        {
            if (!string.IsNullOrEmpty(Pattern.FunctionName))
            {
                return string.IsNullOrEmpty(Pattern.Failflag) ? $"{HardIpConstData.BinFlowFlag}_{BlockName}_{SubBlockName}" : Pattern.Failflag.Replace("F_", "Bin_");
            }

            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""));
        }

        private string CreateRtosOpcode()
        {
            return RtosOpcode;
        }

        private List<string> CreateRtosItems()
        {
            return new List<string> { "T" };
        }
    }
}
