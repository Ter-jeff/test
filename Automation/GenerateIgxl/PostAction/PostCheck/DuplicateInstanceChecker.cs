using System.Linq;

using Automation.Static;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.PostAction.PostCheck
{
    public class DuplicateInstanceChecker
    {
        public void WorkFlow()
        {
            var duplicateInst = TestProgram.IgxlWorkBk.InsSheets.Select(x => x.Value).SelectMany(y => y.Rows).GroupBy(p => p.TestName).Where(p => p.Count() > 1).ToList();
            foreach (IGrouping<string, InstanceRow> group in duplicateInst)
            {
                string testName = group.First().TestName;
                int errorRowNum = group.First().RowNum;
                string duplicateTimes = group.Count().ToString();
                if (!string.IsNullOrEmpty(testName))
                {
                    string sheets = string.Join(",", group.Select(x => x.SheetName).Distinct());
                    ErrorReportManager.AddError(PostActionErrorType.W_DuplicateInstance_01, sheets, errorRowNum, 0,
                        [testName, duplicateTimes, sheets]);
                }
            }
        }
    }
}
