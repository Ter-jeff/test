using System.Collections.Generic;
using System.Linq;

using BinCutScriptLib.Base.Line;

using CommonLib.Datalog;

namespace BinCutScriptLib.Base
{
    public class InstanceData
    {
        public string InstanceName = string.Empty;
        // algorithm chek result
        public bool IsCheckPassByInstance;
        public List<LimitRow> PatternResultRows = [];
        public List<List<PatternRow>> PatternRows = [];
        public List<PatternInfo> FailPatternData = [];
        public int PowersIdx;
        public int FinalStep;
        public double Ids;
        public double Lvcc;
        public int Eqns;
        public int Bin;
        //to study step increment efficiency
        public int UsedSteps;
        //to study step increment efficiency
        public int ActualSteps;

        public string InDatalogKey = "";
        public bool IsSearch;

        public InstanceData() { }

        public InstanceData(InstanceData instanceData)
        {
            if (instanceData == null)
            {
                return;
            }

            InstanceName = instanceData.InstanceName;
            IsCheckPassByInstance = instanceData.IsCheckPassByInstance;
            PatternResultRows = [.. instanceData.PatternResultRows.Select(x => x.Copy())];
            PatternRows = [.. instanceData.PatternRows.Select(rows => rows.Select(x => x.Copy()).ToList())];
            FailPatternData = [.. instanceData.FailPatternData.Select(x => x.Copy())];
            PowersIdx = instanceData.PowersIdx;
            FinalStep = instanceData.FinalStep;
            Ids = instanceData.Ids;
            Lvcc = instanceData.Lvcc;
            Eqns = instanceData.Eqns;
            Bin = instanceData.Bin;
            UsedSteps = instanceData.UsedSteps;
            ActualSteps = instanceData.ActualSteps;
            InDatalogKey = instanceData.InDatalogKey;
            IsSearch = instanceData.IsSearch;
        }

        public InstanceData Copy()
        {
            return new InstanceData(this);
        }
    }
}
