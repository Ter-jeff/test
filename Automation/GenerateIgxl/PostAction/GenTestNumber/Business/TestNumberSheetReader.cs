using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.PostAction.GenTestNumber.Base;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Utility;

using OfficeOpenXml;

namespace Automation.GenerateIgxl.PostAction.GenTestNumber.Business
{
    internal class TestNumberSheetReader
    {
        public Dictionary<string, TestNumberBase> TestNumList = new Dictionary<string, TestNumberBase>(StringComparer.CurrentCultureIgnoreCase);
        public Dictionary<string, (string, string)> UsedTnList = new Dictionary<string, (string, string)>(StringComparer.CurrentCultureIgnoreCase);

        private string CellStr(string inStr)
        {
            inStr = Convert.ToString(inStr).Replace(" ", "");
            inStr = Convert.ToString(inStr).Replace("\n", "");
            inStr = Convert.ToString(inStr).Replace("(", "");
            inStr = Convert.ToString(inStr).Replace(")", "");
            return inStr;
        }

        public TestNumberSheetReader(string workBookPath, string sheetName)
        {
            #region Information
            var filePath = new FileInfo(workBookPath);
            string flowKeyWdSearch = CellStr("Sub Flow Name");
            string numStartWdSearch = CellStr("Testnumber (start)");
            string numEndWdSearch = CellStr("Testnumber (end)");
            string numItvWdSearch = CellStr("Step Num");
            int stRow = -1;
            int spRow = -1;
            int flowCol = -1;
            int numStCol = -1;
            int numSpCol = -1;
            int stepCol = -1;
            ExcelPackage epPlan = new ExcelPackage(filePath);
            ExcelWorksheet ws = epPlan.Workbook.Worksheets[sheetName];
            int searchColMax = 99;
            if (ws.Dimension.End.Column < searchColMax)
            {
                searchColMax = ws.Dimension.End.Column;
            }

            #endregion

            #region To obtain the column locations
            for (int iRow = 1; iRow <= ws.Dimension.End.Row; iRow++)
            {
                for (int iCol = 1; iCol <= searchColMax; iCol++)
                {
                    MatchHeaderColumns(ws, iRow, iCol, flowKeyWdSearch, numStartWdSearch, numEndWdSearch, numItvWdSearch,
                        ref stRow, ref flowCol, ref numStCol, ref numSpCol, ref stepCol);
                    if (flowCol > 0 && numStCol > 0 && stepCol > 0)
                    {
                        break;
                    }
                }
            }

            #endregion

            #region To obtain the Maximum Row
            for (int iRow = stRow; iRow <= ws.Dimension.End.Row; iRow++)
            {
                spRow = iRow;
                if (CellStr(Convert.ToString(ws.Cells[iRow, flowCol].Value)) == "")
                {
                    break;
                }
            }
            #endregion

            #region To obtain the values
            string lastSubFlow = "";
            for (int iRow = stRow + 1; iRow <= spRow; iRow++)
            {
                ReadTestNumberRow(ws, iRow, filePath, flowCol, numStCol, numSpCol, stepCol, ref lastSubFlow);
            }
            #endregion

            #region Check the max test number against next subflow
            string testNumSpErr = CheckTestNumberOverlap();
            #endregion

            #region Error exporting
            if (testNumSpErr.Length > 0)
            {
                ErrorMessageBox.Show(workBookPath + "\n" + testNumSpErr, "TestNumber Assignment Contained Error");
            }
            #endregion
        }

        private void MatchHeaderColumns(ExcelWorksheet ws, int iRow, int iCol, string flowKeyWdSearch, string numStartWdSearch,
            string numEndWdSearch, string numItvWdSearch, ref int stRow, ref int flowCol, ref int numStCol, ref int numSpCol, ref int stepCol)
        {
            if (Regex.IsMatch(CellStr(Convert.ToString(ws.Cells[iRow, iCol].Value)), flowKeyWdSearch, RegexOptions.IgnoreCase))
            {
                stRow = iRow;
                flowCol = iCol;
            }
            if (Regex.IsMatch(CellStr(Convert.ToString(ws.Cells[iRow, iCol].Value)), numStartWdSearch, RegexOptions.IgnoreCase))
            {
                numStCol = iCol;
            }
            if (Regex.IsMatch(CellStr(Convert.ToString(ws.Cells[iRow, iCol].Value)), numEndWdSearch, RegexOptions.IgnoreCase))
            {
                numSpCol = iCol;
            }
            if (Regex.IsMatch(CellStr(Convert.ToString(ws.Cells[iRow, iCol].Value)), numItvWdSearch, RegexOptions.IgnoreCase))
            {
                stepCol = iCol;
            }
        }

        private void ReadTestNumberRow(ExcelWorksheet ws, int iRow, FileInfo filePath, int flowCol, int numStCol, int numSpCol,
            int stepCol, ref string lastSubFlow)
        {
            TestNumberBase newItem = new TestNumberBase();
            long tmp;
            string subFlow = CellStr(Convert.ToString(ws.Cells[iRow, flowCol].Value)).ToUpper().Trim();
            if (subFlow == "")
            {
                return;
            }

            if (CellStr(Convert.ToString(ws.Cells[iRow, numStCol].Value)) != "")
            {
                long.TryParse(CellStr(Convert.ToString(ws.Cells[iRow, numStCol].Value)), out tmp);
                newItem.StartNum = tmp;
                if (lastSubFlow.Length > 0)
                {
                    if (TestNumList[lastSubFlow].MaxNum == 999999999)
                    {
                        TestNumList[lastSubFlow].MaxNum = newItem.StartNum - 1;
                    }
                }
            }
            if (CellStr(Convert.ToString(ws.Cells[iRow, numSpCol].Value)) != "")
            {
                long.TryParse(CellStr(Convert.ToString(ws.Cells[iRow, numSpCol].Value)), out tmp);
                newItem.MaxNum = tmp;
            }
            if (CellStr(Convert.ToString(ws.Cells[iRow, stepCol].Value)) != "")
            {
                long.TryParse(CellStr(Convert.ToString(ws.Cells[iRow, stepCol].Value)), out tmp);
                newItem.Interval = tmp;
            }

            if (!TestNumList.ContainsKey(subFlow))
            {
                TestNumList.Add(subFlow, newItem);
            }
            else
            {
                ErrorReportManager.AddError(BasicErrorType.E_DuplicateItems_01, filePath.Name, 0, 0, $"TN_Assignment sheet contains duplicate sheets : {subFlow}", new string[] { subFlow }, new ErrorInfo() { Comments = new List<string>() { subFlow } });
            }

            lastSubFlow = subFlow;
        }

        private string CheckTestNumberOverlap()
        {
            string testNumSpErr = "";
            int checkCnt = 0;
            long maxLast = 0;
            foreach (string key in TestNumList.Keys)
            {
                if (checkCnt == 0)
                {
                    maxLast = TestNumList[key].MaxNum;
                }
                else
                {
                    if (TestNumList[key].StartNum <= maxLast)
                    {
                        testNumSpErr += "[Error] Subflow=" + key + ", the Start Test# less than last subflow End! Start#=" + TestNumList[key].StartNum + ";  last End#=" + maxLast + "\n";

                    }
                    maxLast = TestNumList[key].MaxNum;
                }
                checkCnt++;
            }
            return testNumSpErr;
        }
    }
}
