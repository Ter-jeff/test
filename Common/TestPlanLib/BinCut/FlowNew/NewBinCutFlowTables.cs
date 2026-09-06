using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using TestPlanLib.BinCut.Binning;

namespace TestPlanLib.BinCut.FlowNew
{
    public partial class NewBinCutFlowTables : List<NewBinCutFlowTable>
    {
        [GeneratedRegex("/s{2,}", RegexOptions.Compiled)]
        private static partial Regex MyRegex();
        private static readonly Regex _regex = MyRegex();

        public List<string> JobList = [];
        public void Check(BinningTables binningTables)
        {
            var modes = new List<string>();
            if (binningTables?.Count > 0)
            {
                modes = [.. binningTables[0].Rows.Select(x => x.RowData[binningTables[0].ModeIdx]).Distinct()];
            }

            for (int i = 0; i < Count; i++)
            {
                #region syntax of mode
                if (modes.Count != 0)
                {
                    for (int j = i + 1; j < this[i].Rows.Count; j++)
                    {
                        NewBinCutFlowSheetRow row = this[i].Rows[j];
                        string name = row.PerformanceMode.Split('_').First();
                        if (!modes.Exists(x => x.EqualsIgnoreCase(name)))
                        {
                            string errorMessage = $"Please check the {row.PerformanceMode} syntax of performance, that should be existed in binning sheeet !!!";
                            this[i].AddError(BinCutErrorType.W_Flow_06, this[i].SheetName, row.RowNum, this[i].PerformanceModeIndex, $"Please check the {row.PerformanceMode} syntax of performance, that should be existed in binning sheeet !!!", [row.PerformanceMode]);
                        }
                    }
                }
                #endregion
            }
        }

        public List<string> GetFlowNames()
        {
            var flowNames = new List<string>();
            foreach (NewBinCutFlowTable newbinCutFlowSheet in this)
            {
                flowNames.AddRange(newbinCutFlowSheet.Rows.SelectMany(x => x.SubFlows));
                flowNames = [.. flowNames.Where(x => x != "0" && !string.IsNullOrEmpty(x)).Distinct().Select(x => x.Trim()).Select(flowName => _regex.Replace(flowName, " ").Replace("#", "[ |_]").Split([':'], StringSplitOptions.RemoveEmptyEntries).Last().Trim())];
            }

            return flowNames;
        }

        //public List<NewBinCutFlowSheetRow> GetRowAt(int rowNum)
        //{
        //    var rows = new List<NewBinCutFlowSheetRow>();
        //    for (var i = 0; i < this.Count(); i++)
        //    {
        //        for (var j = i + 1; j < this[i].Rows.Count(); j++)
        //        {
        //            var row = this[i].Rows[j];
        //            if (row.RowNum == rowNum)
        //                rows.Add(row);
        //        }
        //    }
        //    return rows;
        //}
    }
}
