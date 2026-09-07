using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;

using TestPlanLib.BinNumber;
using TestPlanLib.Singleton;

namespace TestPlanLib.Utility
{
    internal static class VddBinningDefHelpers
    {
        private const string ConBaseVoltage = "Base Voltage";
        private const string ConStepSize = "Step Size";
        private const string ConModeHeader = "Mode";

        private const string ConTd = "TD";
        private const string ConMbist = "Mbist";
        private const string ConSpi = "RTOS";
        private const string ConElb = "ELB";
        private const string ConIlb = "ILB";
        private const string ConBinningFail = "Binning Fail";
        private const string ConLvccFail = "LVCC Fail";

        internal static void AddBaseVoltage(DataTable dataTable, double baseVoltage)
        {
            string err;
            if (string.IsNullOrEmpty(dataTable.Rows[1][0].ToString()!) && string.IsNullOrEmpty(dataTable.Rows[1][1].ToString()!))
            {
                dataTable.Rows[1][0] = ConBaseVoltage;
                dataTable.Rows[1][1] = baseVoltage;
            }
            else if (dataTable.Rows[1][0].ToString()!.Trim().EqualsIgnoreCase(ConBaseVoltage) &&
                string.IsNullOrEmpty(dataTable.Rows[1][1].ToString()!))
            {
                dataTable.Rows[1][1] = baseVoltage;
            }
            else if (dataTable.Rows[1][0].ToString()!.Trim().EqualsIgnoreCase(ConBaseVoltage) &&
                dataTable.Rows[1][1].ToString()!.Trim() != baseVoltage.ToString())
            {
                err = "BinCut Base voltage counflict";
                Response.Report("BinCut has errors : " + err, EnumMessageLevel.Error, 100);
            }
            else if (!dataTable.Rows[1][0].ToString()!.Trim().EqualsIgnoreCase(ConBaseVoltage))
            {
                err = "BinCut Base voltage counflict or mistach";
                Response.Report("BinCut has errors : " + err, EnumMessageLevel.Error, 100);
            }
        }

        internal static void AddStepSize(DataTable dataTable, double stepSize)
        {
            string err;
            if (string.IsNullOrEmpty(dataTable.Rows[1][17].ToString()!) && string.IsNullOrEmpty(dataTable.Rows[1][18].ToString()!))
            {
                dataTable.Rows[1][17] = ConStepSize;
                dataTable.Rows[1][18] = stepSize;
            }
            else if (dataTable.Rows[1][17].ToString()!.Trim().EqualsIgnoreCase(ConStepSize) &&
                string.IsNullOrEmpty(dataTable.Rows[1][18].ToString()!))
            {
                dataTable.Rows[1][18] = stepSize;
            }
            else if (dataTable.Rows[1][17].ToString()!.Trim().EqualsIgnoreCase(ConStepSize) &&
                dataTable.Rows[1][18].ToString()!.Trim() != stepSize.ToString())
            {
                err = "BinCut Step Size counflict";
                Response.Report("BinCut has errors : " + err, EnumMessageLevel.Error, 100);
            }
            else if (!dataTable.Rows[1][17].ToString()!.Trim().EqualsIgnoreCase(ConStepSize))
            {
                err = "BinCut Step Size counflict or mistach";
                Response.Report("BinCut has errors : " + err, EnumMessageLevel.Error, 100);
            }
        }

        internal static void AdjustStartRow(DataTable dataTable)
        {
            int modeRow = FindSpecRow(ConModeHeader, dataTable);
            if (modeRow == 0)
            {
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.InsertAt(newRow, 0);
                dataTable.Rows.InsertAt(newRow, 0);
            }
            if (modeRow == 1)
            {
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.InsertAt(newRow, 1);
            }
        }

        internal static int FindSpecColumn(string pKeyword, DataTable dataTable)
        {
            const int lIResult = 0;
            for (int j = 0; j < 10; j++)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (Regex.IsMatch(dataTable.Rows[j][i].ToString()!, pKeyword, RegexOptions.IgnoreCase))
                    {
                        return i;
                    }
                }
            }
            return lIResult;
        }

        internal static int FindSpecRow(string pKeyword, DataTable dataTable)
        {
            const int lIResult = 0;
            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (Regex.IsMatch(dataTable.Rows[j][i].ToString()!, pKeyword, RegexOptions.IgnoreCase))
                    {
                        return j;
                    }
                }
            }
            return lIResult;
        }

        internal static void RemovedBlankBeforeData(DataTable dataTable, int startRow)
        {
            int modeColumn = FindSpecColumn(ConModeHeader, dataTable);
            for (int i = startRow; i < dataTable.Rows.Count; i++)
            {
                if (dataTable.Rows[i][modeColumn].ToString()!.Length == 0)
                {
                    dataTable.Rows[i].Delete();
                }
                if (dataTable.Rows[i][modeColumn].ToString()!.Length != 0)
                {
                    return;
                }
            }
            dataTable.AcceptChanges();
        }

        internal static void FormatByIsCsharp(bool isCsharp, out List<string> bintypes, out List<string> lvccAndBinnings)
        {
            bintypes = [ConTd, ConMbist];
            lvccAndBinnings = [ConLvccFail];
            if (!isCsharp)
            {
                bintypes.Add(ConSpi);
                lvccAndBinnings.Add(ConBinningFail);
            }
            else
            {
                bintypes.Add(ConElb);
                bintypes.Add(ConIlb);
            }
        }

        internal static void AddBincutBinNums(DataTable dataTable, bool isCsharp, int i, string performanceMode, KeyValuePair<string, int> column, string category1, string category2)
        {
            if (!isCsharp)
            {
                if (BinNumberSingleton.BincutBinNums.ContainsKey(performanceMode + "," + category2))
                {
                    dataTable.Rows[i][column.Value] = BinNumberSingleton.BincutBinNums[performanceMode + "," + category2].BinNumInfo.HardBin.ToString("G15");
                }
                else
                {
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", category1, category2);
                    BinNumberSingleton.BincutBinNums.Add(performanceMode + "," + category2, binNumInfo);
                    dataTable.Rows[i][column.Value] = binNumInfo.BinNumInfo.HardBin.ToString("G15");
                }
            }
            else
            {
                if (!BinNumberSingleton.BincutBinNums.ContainsKey(performanceMode + "," + category2))
                {
                    BinNumResult binNumInfo = BinNumberSingleton.Instance.GetBinInfo("Bincut", category1, category2);
                    BinNumberSingleton.BincutBinNums.Add(performanceMode + "," + category2, binNumInfo);
                }
            }
        }
    }
}
