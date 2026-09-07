using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Reader;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public abstract class ContiBase
    {
        public DcTestContiRow DcTestContiRow;

        public string TestPinGroup { set; get; }
        public DcForceCondition ForceCondition { set; get; }
        public Dictionary<string, string> ContiFailFlags { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string CombinedBinName { get; set; }

        protected ContiBase(DcTestContiRow dcTestContiRow)
        {
            DcTestContiRow = dcTestContiRow;
            TestPinGroup = dcTestContiRow.PinGroup;

            string condition = string.Join(";", dcTestContiRow.GetForceCondition());
            if (dcTestContiRow.TestType == ContiType.OpenShort)
            {
                ForceCondition = new DcForceCondition("I", condition);
            }
            else
            {
                ForceCondition = new DcForceCondition("V", condition);
            }
        }

        internal FlowRow CreateTestFlowRow(string testName, string flagName, string columnA = "")
        {
            var row = new FlowRow { ColumnA = columnA, Opcode = OpCode.Test, Parameter = testName, FailAction = flagName };
            return row;
        }

        protected virtual string CreateContiTestName()
        {
            string testName = DcTestContiRow.InstanceName;
            return testName;
        }


        internal string CreateWalkingZPatternNameOnly(bool includePattern)
        {
            string projectName = LocalSpecs.CurrentProject;

            projectName = projectName.Replace('-', '_');

            return includePattern
                ? $"{projectName}_Pattern_{TestPinGroup}_WalkingZ"
                : $"{projectName}_{TestPinGroup}_WalkingZ";
        }

        internal FlowRow CreateBinTableFlowRow(string binName)
        {
            var binOpen = new FlowRow { Opcode = OpCode.BinTable, Parameter = binName };
            return binOpen;
        }

        internal void SetArgumentFromCondition(ref Function function)
        {
            var interposePrePat = new List<string>();
            foreach (KeyValuePair<string, string> condition in DcTestContiRow.ConditionDict)
            {
                if (condition.Key.Equals("VOL", StringComparison.OrdinalIgnoreCase) ||
                    condition.Key.Equals("VOH", StringComparison.OrdinalIgnoreCase) ||
                    condition.Key.Equals("IOL", StringComparison.OrdinalIgnoreCase) ||
                    condition.Key.Equals("IOH", StringComparison.OrdinalIgnoreCase) ||
                    condition.Key.Equals("HiZ", StringComparison.OrdinalIgnoreCase))
                {
                    interposePrePat.Add($"{DcTestContiRow.PinGroup}:{condition.Key}:{condition.Value}");
                }
                else if (condition.Key.Equals("interposePrePat", StringComparison.OrdinalIgnoreCase))
                {
                    interposePrePat.AddRange(condition.Value.Split(';', ','));
                }
                else
                {
                    function.SetParamValue(condition.Key, condition.Value);
                }
            }
            if (interposePrePat.Any())
            {
                function.SetParamValue("interposePrePat", string.Join(";", interposePrePat));
            }
        }

        protected string GetLevels()
        {
            if (DcTestContiRow.ConditionDict.TryGetValue("Levels", out string value))
            {
                return value;
            }
            return GetDefaultLevels();
        }

        protected virtual string GetDefaultLevels()
        {
            return "Levels_Conti";
        }

        internal static class ScaleHelper
        {
            private static readonly Dictionary<string, string> _scaleMap =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "uA", "u" },
                    { "uV", "u" }
                };

            internal static string GetHighScale(string limitUnit)
            {

                if (string.IsNullOrEmpty(limitUnit))
                {
                    return "m";
                }

                return _scaleMap.TryGetValue(limitUnit, out string scale)
                    ? scale
                    : "m";
            }
        }

        internal string SetCategoryFromCondition()
        {
            string dcCategory = "Conti_X_X_X";
            if (DcTestContiRow.ConditionDict.ContainsKey("DC"))
            {
                dcCategory = DcTestContiRow.ConditionDict["DC"];
            }

            return dcCategory;
        }
        public abstract List<FlowRow> GenerateFlowRows();
        public abstract List<InstanceRow> GenerateInstanceRows();
    }
}
