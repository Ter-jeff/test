using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.Scan.Harvest.Flow;
using Automation.InputManager.Data;
using Automation.Utility.HardIP;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow
{
    public class IdsFlowRowGenerator : HardIpFlowRowGenerator
    {
        #region properties
        protected bool IdsNoFuse
        {
            get
            {
                return HardIpService.IdsNoFuse(Pat.MiscInfo);
            }
        }

        protected bool IsNandPattern
        {
            get
            {
                return HardIpService.IsNandPattern(Pat.Pattern.GetLastPayload());
            }
        }


        protected bool IsSpiPattern
        {
            get
            {
                return HardIpService.IsSpiPattern(Pat.Pattern.GetLastPayload());
            }
        }

        #endregion

        #region Constructor
        public IdsFlowRowGenerator(HardIpInputData hardIpInputData, string sheetName) : base(hardIpInputData, sheetName)
        {
        }
        #endregion

        #region Main Methods

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
        }

        public override List<FlowRow> GenTestRows(bool isCz2Only = false)
        {
            var flowRows = new List<FlowRow>();
            var testFlowRow = new FlowRow();
            var hardipMainNonBincutInstance = new HardipNonBinCutInstance();

            testFlowRow.Job = CreateTestJob();
            testFlowRow.Part = CreateTestPart();
            testFlowRow.Opcode = CreateTestOpcode();
            testFlowRow.Env = CreateTestEnv();
            testFlowRow.Enable = GenEnable(CreateEnableWord(), testFlowRow.Env);
            testFlowRow.Parameter = CreateTestParameter();
            testFlowRow.FailAction = CreateTestFailAction();

            //use limit rows
            List<FlowRow> limitRows = null;
            if (IdsNoFuse || IsNandPattern || IsSpiPattern)
            {
                limitRows = limitRows = GenUseLimits(
                    testFlowRow.Parameter,
                    Pat,
                    LabelVoltage,
                    CreateUseLimitFailAction
                );
            }

            List<FlowRow> ifFlowRows = hardipMainNonBincutInstance.GetIfFlowRows(null, [testFlowRow], limitRows);
            flowRows.AddRange(ifFlowRows);

            return flowRows;
        }

        public override List<FlowRow> GenShmooRows(string labelVoltage = "")
        {
            var flowrows = new List<FlowRow>();
            return flowrows;
        }

        public override FlowRow GenBinTableRow(string voltage = "")
        {
            var bintableRow = new FlowRow
            {
                Job = CreateBinTableJob(),
                Opcode = CreateBinTableOpcode(),
                Env = CreateBinTableEnv(),
                Enable = CreateBinTableEnable(),
                Parameter = CreateBinTableParameter()
            };
            return bintableRow;
        }

        #endregion

        #region Create flow row columns methods
        protected string CreateEnableWord()
        {
            return CommonGenerator.GenEnableWord(Pat.Pattern.GetLastPayload(), Pat.MiscInfo, LabelVoltage, HardIpInputData);
        }

        protected override string CreateBinTableParameter()
        {
            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName);
        }

        protected internal override string CreateTestParameter(bool forShmoo = false)
        {
            string patternName = Pat.Pattern.GetPatternName();

            if (!string.IsNullOrEmpty(Pat.TestName))
            {
                return Pat.TestName + "_" + ActualLabelVoltage;
            }

            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, patternName,
                Pat.PatternIndexFlag,
                TimingAc, Pat.ForceVoltageFlag, InstNameSubStr, "NV", NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
        }

        protected internal override string CreateTestFailAction()
        {
            if (Pat.IsIgnorePatBinOut())
            {
                return "";
            }

            return CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), Pat.Pattern.GetLastPayload(), TimingAc,
                InstNameSubStr, "", Pat.MiscInfo, NoPattern);
        }

        protected string CreateUseLimitFailAction(string repeatStr)
        {
            if (Pat.IsIgnorePatBinOut())
            {
                return "";
            }

            if (IdsNoFuse)
            {
                return CommonGenerator.GenHardIpFlowFailAction(SheetName, BlockName, SubBlockName.ToUpper().Replace("-MERGE", ""), Pat.Pattern.GetLastPayload(), TimingAc,
                    InstNameSubStr, "", Pat.MiscInfo, NoPattern);
            }

            if (repeatStr != "")
            {
                return "F_IDS_Current_Main" + "_" + repeatStr.Replace(".", "p");
            }

            return "F_IDS_Current_Main_1x";
        }
        #endregion
    }
}
