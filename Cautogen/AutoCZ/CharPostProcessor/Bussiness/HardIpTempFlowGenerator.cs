using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using FlowStep = Teradyne.Oasis.IGData.FlowStep;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class HardIpTempFlowGenerator
    {
        public static SubFlowSheet GetHardIpTempSubFlow(CharPlanSheet charPlanSheet, Dictionary<string, InstanceRow> testInstancesDict)
        {
            /* retrieve temp flow sheet target hip sheets */
            var tempFlowSheet = new SubFlowSheet("Flow_Temp_" + charPlanSheet.SheetName);
            if (!charPlanSheet.IsHardIp)
            {
                return tempFlowSheet;
            }

            List<string> flowSheets = LocalSpecs.ProgFlowDic[charPlanSheet.SheetName];
            IEnumerable<FlowRow> progFlowSteps =
                LocalSpecs.ProgInfo.AllFlowSteps.Where(x => !string.IsNullOrEmpty(x.SheetName) && flowSheets.Contains(x.SheetName));

            foreach (FlowRow flowRow in
                from flowStep in progFlowSteps
                where flowStep.Opcode.ToLower() != "return"
                where flowStep.Opcode.ToLower() != "for"
                where flowStep.Opcode.ToLower() != "next"
                where !flowStep.Parameter.ToLower().Contains("set_all_ids_data_to_efuse")
                where (!flowStep.SheetName.EqualsIgnoreCase("Flow_DCTEST_IDS")
                    || CheckInstByFunction(flowStep.Parameter, "ControlRelay", testInstancesDict))
                select flowStep)
            {
                tempFlowSheet.AddRow(flowRow);
            }
            return tempFlowSheet;
        }

        private static bool CheckInstByFunction(string instanceName, string functionName, Dictionary<string, InstanceRow> instDict)
        {
            if (instanceName.StartsWith("IDS_", System.StringComparison.OrdinalIgnoreCase)
                && instDict.TryGetValue(instanceName, out InstanceRow instanceRow))
            {
                if (instanceRow.VbtName.Split('.').Last().EqualsIgnoreCase(functionName))
                {
                    return true;
                }
            }
            return false;
        }

        public static FlowRow ConvertRow(FlowStep step)
        {
            return new FlowRow
            {
                Parameter = step.Parameter,
                Label = step.Label,
                Job = string.Join(",", step.GateJob),
                Part = string.Join(",", step.GatePart),
                Opcode = step.Opcode,
                Env = string.Join(",", step.GateEnv),
                HiLim = step.HiLim,
                LoLim = step.LoLim,
                Enable = string.Join(",", step.EnableWords),
                FailAction = step.FailAction
            };
        }
    }
}
