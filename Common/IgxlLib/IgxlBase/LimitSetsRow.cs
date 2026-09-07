using System.Collections.Generic;
using System.Diagnostics;

namespace IgxlLib.IgxlBase
{
    [DebuggerDisplay("{TestName} ")]
    public class LimitSetsRow : IgxlRow
    {
        public string Fields = string.Empty;
        public string Row = string.Empty;
        public string FlowTable = string.Empty;
        public string Instance = string.Empty;
        public string TestName = string.Empty;
        public string TestNumber = string.Empty;
        public string LowCompSign = ">=";
        public string HighCompSign = "<=";
        public List<LimitSetsItem> LimitSetsCat = [];
        public string Scale = string.Empty;
        public string Units = string.Empty;
        public string Format = string.Empty;
        public string PassSort = string.Empty;
        public string FailSort = string.Empty;
        public string SortName = string.Empty;
        public string PassBin = string.Empty;
        public string FailBin = string.Empty;
        public string BinName = string.Empty;
        public string Result = "none";
        public string PassAction = string.Empty;
        public string FailAction = string.Empty;
        public string Assume = string.Empty;
        public string Sites = string.Empty;
        public string Comment = string.Empty;

        public LimitSetsRow()
        {
        }

        public LimitSetsRow(string row, string flowTable, string instance, string testName, string testNumber, List<LimitSetsItem> limitSetsItems, string scale, string units, string format, string passSort, string failSort, string passBin, string failBin, string passAction, string failAction, string assume, string sites, string comment)
        {
            Row = row;
            FlowTable = flowTable;
            Instance = instance;
            TestName = testName;
            TestNumber = testNumber;
            LimitSetsCat = limitSetsItems;
            Scale = scale;
            Units = units;
            Format = format;
            PassSort = passSort;
            FailSort = failSort;
            PassBin = passBin;
            FailBin = failBin;
            PassAction = passAction;
            FailAction = failAction;
            Assume = assume;
            Sites = sites;
            Comment = comment;
        }
    }
}
