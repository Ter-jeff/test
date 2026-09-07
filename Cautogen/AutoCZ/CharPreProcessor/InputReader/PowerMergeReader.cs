using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.Utility;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.InputReader
{
    public class PowerMergeReader
    {
        public static void Read(ExcelWorksheet wSheet)
        {
            UtilityMain.UtilityData.PowerMergeResult.Columns.Add("Master", Type.GetType("System.String"));
            UtilityMain.UtilityData.PowerMergeResult.Columns.Add("FT", Type.GetType("System.String"));
            UtilityMain.UtilityData.PowerMergeResult.Columns.Add("CP", Type.GetType("System.String"));
            int ftNetCol = 1, ftBallCol = 1, wsNetCol = 1, wsBumpCol = 1;
            var master = new List<string>();
            int rowCount = -1;
            ResolveHeaderColumns(wSheet, ref ftNetCol, ref ftBallCol, ref wsNetCol, ref wsBumpCol);

            for (int i = 3; i <= wSheet.Dimension.End.Row; i++)
            {
                string ftNetValue = GetMegerValue(wSheet, i, ftNetCol);
                string ftBallValue = GetMegerValue(wSheet, i, ftBallCol);
                string wsNetValue = GetMegerValue(wSheet, i, wsNetCol);
                string wsBumpValue = GetMegerValue(wSheet, i, wsBumpCol);
                if (!ProcessFtNetRow(master, ftNetValue, ftBallValue, wsNetValue, wsBumpValue, ref rowCount))
                {
                    continue;
                }
                ProcessWsNetRow(master, ftNetValue, ftBallValue, wsNetValue, wsBumpValue, ref rowCount);
            }
        }

        private static void ResolveHeaderColumns(ExcelWorksheet wSheet, ref int ftNetCol, ref int ftBallCol, ref int wsNetCol, ref int wsBumpCol)
        {
            for (int i = 1; i <= wSheet.Dimension.End.Column; i++)
            {
                if (wSheet.Cells[2, i].Value == null)
                {
                    continue;
                }

                string header = wSheet.Cells[2, i].Value.ToString().ToUpper();

                if (Regex.IsMatch(header, @"FT\s*NET", RegexOptions.IgnoreCase))
                {
                    ftNetCol = i;
                }
                else if (Regex.IsMatch(header, @"FT\s*BALL", RegexOptions.IgnoreCase))
                {
                    ftBallCol = i;
                }
                else if (Regex.IsMatch(header, @"WS\s*NET", RegexOptions.IgnoreCase))
                {
                    wsNetCol = i;
                }
                else if (Regex.IsMatch(header, @"WS\s*BUMP", RegexOptions.IgnoreCase))
                {
                    wsBumpCol = i;
                }
            }
        }

        // Returns false to signal `continue` (skip remainder of outer-loop iteration).
        private static bool ProcessFtNetRow(List<string> master, string ftNetValue, string ftBallValue, string wsNetValue, string wsBumpValue, ref int rowCount)
        {
            if (!master.Contains(ftNetValue))
            {
                if (ftNetValue == "N/A")
                {
                    return false;
                }

                master.Add(ftNetValue);
                System.Data.DataRow newRow = UtilityMain.UtilityData.PowerMergeResult.NewRow();
                rowCount++;
                UtilityMain.UtilityData.PowerMergeResult.Rows.Add(newRow);
                UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][0] = ftNetValue;
                if (ftBallValue != "" && ftBallValue != "N/A" && ftBallValue != "VSS")
                {
                    UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][1] = ftNetValue;
                }

                if (wsBumpValue != "" && wsBumpValue != "N/A" && wsBumpValue != "VSS")
                {
                    UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][2] = wsNetValue;
                }
            }
            else if (wsNetValue != "N/A" && wsNetValue != "" && wsBumpValue != "" && wsBumpValue != "N/A" && wsBumpValue != "VSS" && UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][2].ToString() == "")
            {
                UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][2] = wsNetValue;
            }
            return true;
        }

        // Returns false to signal `continue` (skip remainder of outer-loop iteration).
        private static bool ProcessWsNetRow(List<string> master, string ftNetValue, string ftBallValue, string wsNetValue, string wsBumpValue, ref int rowCount)
        {
            if (!master.Contains(wsNetValue))
            {
                if (wsNetValue == "N/A")
                {
                    return false;
                }

                master.Add(wsNetValue);
                System.Data.DataRow newRow = UtilityMain.UtilityData.PowerMergeResult.NewRow();
                rowCount++;
                UtilityMain.UtilityData.PowerMergeResult.Rows.Add(newRow);
                UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][0] = wsNetValue;
                if (ftBallValue != "" && ftBallValue != "N/A" && ftBallValue != "VSS")
                {
                    UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][1] = ftNetValue;
                }

                if (wsBumpValue != "" && wsBumpValue != "N/A" && wsBumpValue != "VSS")
                {
                    UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][2] = wsNetValue;
                }
            }
            else if (ftNetValue != "N/A" && ftNetValue != "" && ftBallValue != "" && ftBallValue != "N/A" && ftBallValue != "VSS" && UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][1].ToString() == "")
            {
                UtilityMain.UtilityData.PowerMergeResult.Rows[rowCount][1] = ftNetValue;
            }
            return true;
        }

        /// <summary> Added by Jackie
        /// Get Megered value with cell address
        /// </summary>
        private static string GetMegerValue(ExcelWorksheet wSheet, int row, int column)
        {
            string range = wSheet.MergedCells[row, column];
            if (range == null)
            {
                return wSheet.Cells[row, column].Value != null ? wSheet.Cells[row, column].Value.ToString() : "";
            }

            object value = wSheet.Cells[new ExcelAddress(range).Start.Row, new ExcelAddress(range).Start.Column].Value;
            return value != null ? value.ToString() : "";
        }
    }
}
