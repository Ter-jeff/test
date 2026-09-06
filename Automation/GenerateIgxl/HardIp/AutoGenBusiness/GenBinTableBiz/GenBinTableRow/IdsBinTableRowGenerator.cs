using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Utility.HardIP;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow
{
    public class IdsBinTableRowGenerator : BinTableRowGeneratorBase
    {
        private const string IdsOpcode = "OR";

        private bool _idsNoFuse;
        private string _repeatStr;
        public IdsBinTableRowGenerator(string sheetName, List<string> errorBinNums)
            : base(sheetName, errorBinNums)
        {

        }

        protected override BinTableRow GenBinTableRowForPattern(string voltage = "")
        {
            var binTableRow = new BinTableRow
            {
                ItemList = CreateIdsItemList(),
                Name = CreateIdsName(),
                Items = CreateIdsItems(),
                Op = CreateIdsOpcode(),
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

        private string CreateIdsItemList()
        {
            if (_idsNoFuse)
            {
                return CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), Pattern.Pattern.GetLastPayload(), TimingAc,
                    InstNameSubStr, "", Pattern.MiscInfo, NoPattern);
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

        private string CreateIdsName()
        {
            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName);
        }

        private string CreateIdsOpcode()
        {
            return IdsOpcode;
        }

        private List<string> CreateIdsItems()
        {
            return Enumerable.Repeat("T", _repeatStr.Split(',').Length).ToList();
        }
    }
}
