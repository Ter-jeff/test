using System;
using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using DebugPlanReaderLib.DebugPlan;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateTmpsFlow
    {
        public List<SubFlowSheet> Work(SubFlowSheet TMPS_Flow, DebugPlanMain debugTestPlan, string job)
        {
            var tmpsFlowSheets = new List<SubFlowSheet>();
            var tmpsGroups = debugTestPlan.AiTestPlanSheets.SelectMany(p => p.Rows)
                .Where(p => !p.GetTmpsFlowName().Equals("Flow_TMPS", StringComparison.OrdinalIgnoreCase))
                .GroupBy(p => p.GetTmpsFlowName()).ToDictionary(p => p.Key, p => p.ToList());
            foreach (var tmpsGroup in tmpsGroups)
            {
                var newTmps_Flow = new SubFlowSheet(tmpsGroup.Key);
                var lowLimit = "";
                var highLimit = "";
                DataConvertor.GetTMPS_temperature(tmpsGroup.Value.First().TempCondition, out lowLimit, out highLimit);
                foreach (var row in TMPS_Flow.Rows)
                {
                    if (row.Units.Equals("C", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(row.Job) ||
                            row.Job.Split(',').ToList().Exists(p => p.Equals(job, StringComparison.OrdinalIgnoreCase)))
                        {
                            var newrow = (FlowRow)(row.Copy());
                            newrow.LoLim = lowLimit;
                            newrow.HiLim = highLimit;
                            newTmps_Flow.AddRow(newrow);
                            continue;
                        }
                    }
                    newTmps_Flow.AddRow(row);
                }
                tmpsFlowSheets.Add(newTmps_Flow);
            }
            return tmpsFlowSheets;
        }
    }
}
