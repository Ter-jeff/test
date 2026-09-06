using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;
using Automation.Static;

using IgxlLib.IgxlBase;

using TestPlanLib.BinCut.Binning;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutFlowInstanceWriterCs : BinCutFlowInstanceWriter
    {
        private readonly BinCutInputData _binCutInputManager;

        public BinCutFlowInstanceWriterCs(BinCutInputData binCutInputManager, List<BinCutFinalInstanceRow> binCutFinalInstanceRows, List<string> gradeSearchJobs = null) : base(binCutInputManager, binCutFinalInstanceRows, gradeSearchJobs)
        {
            _binCutInputManager = binCutInputManager;
        }

        internal override InstanceRow GenPrintOutVddBinning()
        {
            var instanceRow = new InstanceRow();
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PrintVddBinning", "bincut", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenPrintOutVddBinning();
            }

            instanceRow.TestName = "PrintOutVddBinning";
            instanceRow.VbtType = ".NET";
            instanceRow.VbtName = function.FullFunctionName;

            return instanceRow;
        }

        internal override InstanceRow GenAdjustVddBinningInstance(string adjustVddbinning)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("AdjustVddBinning", "bincut", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenAdjustVddBinningInstance(adjustVddbinning);
            }
            if (BinCutInputData.BinningTables.Any())
            {
                VddBinDefComment vddBinComment = BinCutInputData.BinningTables.First().GetAdjustPowerList();
                if (vddBinComment != null)
                {
                    if (vddBinComment.IsMax && vddBinComment.MaxStr != null)
                    {
                        function.SetParamValue("binXExpansionMode", "2");
                        var instanceRow = new InstanceRow
                        {
                            TestName = adjustVddbinning,
                            VbtType = ".NET",
                            VbtName = function.FullFunctionName,
                            ArgList = function.Parameters,
                            Args = function.ArgList
                        };
                        return instanceRow;
                    }

                    if (vddBinComment.IsMin && vddBinComment.MinStr != null)
                    {
                        var instanceRow = new InstanceRow
                        {
                            TestName = adjustVddbinning,
                            VbtType = ".NET",
                            VbtName = function.FullFunctionName,
                            ArgList = function.Parameters,
                            Args = function.ArgList
                        };
                        return instanceRow;
                    }
                    else
                    {
                        var instanceRow = new InstanceRow { TestName = adjustVddbinning, VbtType = ".NET", VbtName = function.FullFunctionName };
                        if (!string.IsNullOrEmpty(function.Parameters))
                        {
                            instanceRow.ArgList = function.Parameters;
                        }

                        return instanceRow;
                    }
                }
            }

            var instanceRow1 = new InstanceRow { TestName = adjustVddbinning, VbtType = ".NET", VbtName = function.FullFunctionName };
            if (!string.IsNullOrEmpty(function.Parameters))
            {
                instanceRow1.ArgList = function.Parameters;
            }

            return instanceRow1;
        }

        internal override InstanceRow GenCheck_IDS()
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("Judge_Stored_IDS", "bincut", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenCheck_IDS();
            }

            var instanceRow = new InstanceRow { TestName = "Judge_stored_IDS", VbtType = ".NET" };
            function.SetParamValue("instanceFailFlag", "F_Vddbinning_IDS_fail");
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;
            return instanceRow;
        }

        internal override InstanceRow GenPower_Binning_Calculation()
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PowerBinning", "bincut", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenPower_Binning_Calculation();
            }

            var instanceRow = new InstanceRow
            {
                TestName = "Power_Binning",
                VbtType = ".NET",
                VbtName = function.FullFunctionName,
                ArgList = function.Parameters,
                Args = function.ArgList
            };
            return instanceRow;
        }

        internal override InstanceRow GenPrintConfigInstanceRow(IEnumerable<string> harvestFlags)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("PrintOutBinCutConfig", "bincut", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenPrintConfigInstanceRow(harvestFlags);
            }

            var instanceRow = new InstanceRow { TestName = "Print_BinCut_config", VbtType = ".NET" };
            function.SetParamValue("otherFlags", string.Join(",", harvestFlags));
            instanceRow.VbtName = function.FullFunctionName;
            instanceRow.ArgList = function.Parameters;
            instanceRow.Args = function.ArgList;

            return instanceRow;
        }

        internal override InstanceRow GenSet_VBinResult_without_Test(Dictionary<string, bool> flowNameList, Dictionary<string, BinCutExtraPolationModeInfo> binCutExtraPolationModeList = null)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("SetVoltageWithoutTest", "bincut", true);
            BinningTable binningTable = _binCutInputManager.BinningTables.FirstOrDefault();
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenSet_VBinResult_without_Test(flowNameList);
            }

            var instanceRow = new InstanceRow
            {
                TestName = "Set_VBinResult_without_Test",
                VbtType = ".NET",
                VbtName = function.FullFunctionName,
                ArgList = function.Parameters,
                Args = function.ArgList
            };
            if (binningTable != null)
            {
                List<string> allowEqualModes = binningTable.GetAllowEqualModes();
                var allowEqualModesWithoutSearch = allowEqualModes.Where(x => !flowNameList.Select(y => y.Key).Contains(x) || !flowNameList[x]).ToList();
                allowEqualModesWithoutSearch.InsertRange(0, binCutExtraPolationModeList.Where(x => x.Value.IsOnlyExtraMode).Select(x => x.Key).ToList());
                instanceRow.SetArgument("allowEqualModes", string.Join(",", allowEqualModesWithoutSearch));
                instanceRow.SetArgument("singleEquationModes", string.Join(",", binningTable.GetSingleEquationModes()));
            }
            return instanceRow;
        }

#nullable enable
        internal override InstanceRow? GenFuseBinnedProductVoltagesInstanceRow()
#nullable restore
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("FuseBinCutResults", "bincut", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenFuseBinnedProductVoltagesInstanceRow();
            }

            var instanceRow = new InstanceRow
            {
                TestName = "FuseBinnedProductVoltages",
                VbtType = ".NET",
                VbtName = function.FullFunctionName,
                ArgList = function.Parameters,
                Args = function.ArgList
            };

            return instanceRow;
        }
    }
}
