using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{Parameter} ")]
    public class FlowRow : IgxlRow
    {
        public string Label { get; set; } = string.Empty;
        public string Enable { get; set; } = string.Empty;
        public string Job { get; set; } = string.Empty;
        public string Part { get; set; } = string.Empty;
        public string Env { get; set; } = string.Empty;
        public string Opcode { get; set; } = string.Empty;
        public string Parameter { get; set; } = string.Empty;
        public string TName { get; set; } = string.Empty;
        public string TNum { get; set; } = string.Empty;
        public string LoLim { get; set; } = string.Empty;
        public string HiLim { get; set; } = string.Empty;
        public string Scale { get; set; } = string.Empty;
        public string Units { get; set; } = string.Empty;
        public string Format { get; set; } = string.Empty;
        public string BinPass { get; set; } = string.Empty;
        public string BinFail { get; set; } = string.Empty;
        public string SortPass { get; set; } = string.Empty;
        public string SortFail { get; set; } = string.Empty;
        public string Result { get; set; } = string.Empty;
        public string PassAction { get; set; } = string.Empty;
        public string FailAction { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string GroupSpecifier { get; set; } = string.Empty;
        public string GroupSense { get; set; } = string.Empty;
        public string GroupCondition { get; set; } = string.Empty;
        public string GroupName { get; set; } = string.Empty;
        public string DeviceSense { get; set; } = string.Empty;
        public string DeviceCondition { get; set; } = string.Empty;
        public string DeviceName { get; set; } = string.Empty;
        public string DebugAssume { get; set; } = string.Empty;
        public string DebugSites { get; set; } = string.Empty;
        public string CtProfileDataElapsedTime { get; set; } = string.Empty;
        public string CtProfileDataBackgroundType { get; set; } = string.Empty;
        public string CtProfileDataSerialize { get; set; } = string.Empty;
        public string CtProfileDataResourceLock { get; set; } = string.Empty;
        public string CtProfileDataFlowStepLocked { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string Comment1 { get; set; } = string.Empty;
        public bool IsSsn { get; set; }
        public string PatName { get; set; } = string.Empty;
        public int PatIndex { get; set; } = -1;
        public int BinTableFlagCount { get; set; }

        public FlowRow()
        {
        }

        public FlowRow(FlowRow flowRow) : base(flowRow)
        {
            if (flowRow == null)
            {
                return;
            }
            ColumnA = flowRow.ColumnA;
            Label = flowRow.Label;
            Enable = flowRow.Enable;
            Job = flowRow.Job;
            Part = flowRow.Part;
            Env = flowRow.Env;
            Opcode = flowRow.Opcode;
            Parameter = flowRow.Parameter;
            TName = flowRow.TName;
            TNum = flowRow.TNum;
            LoLim = flowRow.LoLim;
            HiLim = flowRow.HiLim;
            Scale = flowRow.Scale;
            Units = flowRow.Units;
            Format = flowRow.Format;
            BinPass = flowRow.BinPass;
            BinFail = flowRow.BinFail;
            SortPass = flowRow.SortPass;
            SortFail = flowRow.SortFail;
            Result = flowRow.Result;
            PassAction = flowRow.PassAction;
            FailAction = flowRow.FailAction;
            State = flowRow.State;
            GroupSpecifier = flowRow.GroupSpecifier;
            GroupSense = flowRow.GroupSense;
            GroupCondition = flowRow.GroupCondition;
            GroupName = flowRow.GroupName;
            DeviceSense = flowRow.DeviceSense;
            DeviceCondition = flowRow.DeviceCondition;
            DeviceName = flowRow.DeviceName;
            DebugAssume = flowRow.DebugAssume;
            DebugSites = flowRow.DebugSites;
            CtProfileDataElapsedTime = flowRow.CtProfileDataElapsedTime;
            CtProfileDataBackgroundType = flowRow.CtProfileDataBackgroundType;
            CtProfileDataSerialize = flowRow.CtProfileDataSerialize;
            CtProfileDataResourceLock = flowRow.CtProfileDataResourceLock;
            CtProfileDataFlowStepLocked = flowRow.CtProfileDataFlowStepLocked;
            Comment = flowRow.Comment;
            Comment1 = flowRow.Comment1;
            IsSsn = flowRow.IsSsn;
            PatName = flowRow.PatName;
            PatIndex = flowRow.PatIndex;
            BinTableFlagCount = flowRow.BinTableFlagCount;
        }

        public FlowRow Copy()
        {
            return new FlowRow(this);
        }

        public List<string> GetJobs()
        {
            List<string> jobs = [.. Job.Split(',').Where(x => !string.IsNullOrEmpty(x))];
            return jobs;
        }

        public static FlowRow GenEndIf(string job)
        {
            return new FlowRow { Opcode = "EndIf" };
        }

        public static FlowRow GenIfCondition(string condition, string job)
        {
            return new FlowRow { Opcode = "If", Parameter = "(" + condition + ")" + " || F_Debug_all" };
        }

        public void AddColumnA(string format)
        {
            if (string.IsNullOrEmpty(ColumnA))
            {
                ColumnA = format;
            }
            else
            {
                ColumnA = ColumnA + "," + format;
            }
        }
    }
}
