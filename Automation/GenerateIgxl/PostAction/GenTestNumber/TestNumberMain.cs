using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.GenerateIgxl.PostAction.GenTestNumber.Business;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.PostAction.GenTestNumber
{
    public class TestNumberMain
    {
        private readonly string _settingFilesTnAssignment;
        private const string Nsheet = "Block_TestNumber";
        private List<string> _doneSheets = new List<string>();

        public TestNumberMain(string settingFilesTnAssignment)
        {
            _settingFilesTnAssignment = settingFilesTnAssignment.Replace(Directory.GetCurrentDirectory(), "");
        }

        public List<string> WorkFlow(TestNumberParaData testNumberParaData)
        {
            var dicTNmapping = new TestNumberSheetReader(testNumberParaData.TnXlsx, Nsheet);
            Dictionary<string, SubFlowSheet> subflowSheets = TestProgram.IgxlWorkBk.SubFlowSheets;
            var nonTNsheets = new List<string>();
            Dictionary<string, MainFlow> mainFlowSheets = TestProgram.IgxlWorkBk.MainFlowSheets;
            foreach (KeyValuePair<string, SubFlowSheet> subflowSheet in subflowSheets)
            {
                string subflowName = subflowSheet.Value.Name;
                if (dicTNmapping.TestNumList.ContainsKey(subflowName) ||
                    (!dicTNmapping.TestNumList.ContainsKey(subflowName) && dicTNmapping.TestNumList.ContainsKey(subflowSheet.Value.SplitFromSheet)))
                {
                    var rows = new FlowRows();
                    rows.AddRange(subflowSheet.Value.Rows);

                    string skeetKeyInTestNumList = dicTNmapping.TestNumList.ContainsKey(subflowName) ? subflowName : subflowSheet.Value.SplitFromSheet;
                    long startNumber = dicTNmapping.TestNumList[skeetKeyInTestNumList].StartNum;
                    long step = dicTNmapping.TestNumList[skeetKeyInTestNumList].Interval;
                    int currCnt = dicTNmapping.TestNumList[skeetKeyInTestNumList].CurrCnt;
                    long maxNum = dicTNmapping.TestNumList[skeetKeyInTestNumList].MaxNum;

                    _doneSheets.Add(subflowName);
                    rows.SetTestNumber(TestProgram.IgxlWorkBk.SubFlowSheets, startNumber, step, maxNum, ref currCnt, ref _doneSheets);
                    dicTNmapping.TestNumList[skeetKeyInTestNumList].CurrCnt = currCnt;

                    foreach (FlowRow row in rows.Where(x => !string.IsNullOrEmpty(x.TNum)))
                    {
                        if (dicTNmapping.UsedTnList.TryGetValue(row.TNum, out (string, string) flowAndInstanceName))
                        {
                            ErrorReportManager.AddError(PostActionErrorType.W_DuplicateTestNumber_01, _settingFilesTnAssignment, 0, 0,
                                 [row.TNum, subflowName, row.Parameter, flowAndInstanceName.Item1, flowAndInstanceName.Item2]);
                        }
                        else
                        {
                            dicTNmapping.UsedTnList.Add(row.TNum, (subflowName, row.Parameter));
                        }
                    }
                }
                else
                {
                    if (subflowSheet.Value.Rows.Any(x => x.Opcode.Equals(OpCode.Test, StringComparison.CurrentCultureIgnoreCase) ||
                        x.Opcode.Equals("test-defer-limits", StringComparison.CurrentCultureIgnoreCase)))
                    {
                        nonTNsheets.Add(subflowName);
                    }
                }
            }
            foreach (KeyValuePair<string, MainFlow> mainFlowSheet in mainFlowSheets)
            {
                string mainflowName = mainFlowSheet.Value.Name;
                if (dicTNmapping.TestNumList.ContainsKey(mainflowName))
                {
                    var rows = new FlowRows();
                    rows.AddRange(mainFlowSheet.Value.Rows);
                    int count = 0;
                    long startNumber = dicTNmapping.TestNumList[mainflowName].StartNum;
                    long step = dicTNmapping.TestNumList[mainflowName].Interval;
                    long maxNum = dicTNmapping.TestNumList[mainflowName].MaxNum;
                    _doneSheets.Add(mainflowName);
                    Dictionary<string, SubFlowSheet> subFlowDic = TestProgram.IgxlWorkBk.MainFlowSheets
                        .Where(kvp => kvp.Value is SubFlowSheet)
                        .ToDictionary(
                            kvp => kvp.Key,
                            kvp => (SubFlowSheet)kvp.Value,
                            StringComparer.OrdinalIgnoreCase // Matches your IgnoreCase setting
                        );
                    rows.SetTestNumber(subFlowDic, startNumber, step, maxNum, ref count, ref _doneSheets);
                    foreach (FlowRow row in rows.Where(x => !string.IsNullOrEmpty(x.TNum)))
                    {
                        if (dicTNmapping.UsedTnList.TryGetValue(row.TNum, out (string, string) flowAndInstanceName))
                        {
                            ErrorReportManager.AddError(PostActionErrorType.W_DuplicateTestNumber_01, _settingFilesTnAssignment, 0, 0,
                                [row.TNum, mainflowName, row.Parameter, flowAndInstanceName.Item1, flowAndInstanceName.Item2]);
                        }
                        else
                        {
                            dicTNmapping.UsedTnList.Add(row.TNum, (mainflowName, row.Parameter));
                        }
                    }
                }
                else
                {
                    if (mainFlowSheet.Value.Rows.Any(x => x.Opcode.Equals(OpCode.Test, StringComparison.CurrentCultureIgnoreCase) ||
                        x.Opcode.Equals("test-defer-limits", StringComparison.CurrentCultureIgnoreCase)))
                    {
                        nonTNsheets.Add(mainflowName);
                    }
                }
            }
            return nonTNsheets.Except(_doneSheets, StringComparer.CurrentCultureIgnoreCase).ToList();
        }
    }
}
