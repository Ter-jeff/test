using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using CommonLib.Extension;

using OfficeOpenXml;

namespace TestPlanLib.PowerMerge
{
    public class PowerMergeReader
    {
        /// <summary>
        /// GetPowerMergeDataSet
        /// </summary>
        /// <param name="excelWorksheet"></param>
        /// <returns></returns>
        public static PowerMergeSheet? ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet == null)
            {
                return null;
            }

            ExcelWorksheet lObjSheetAssignment = excelWorksheet;

            if (!FindFirstDataLocation(lObjSheetAssignment, out int lIStartRowIndex, out int lIStartColumnIndex))
            {
                // can not find the first data
                return null;
            }

            Hashtable lPowerTableList = GetPowerTableList(lObjSheetAssignment, lIStartRowIndex, lIStartColumnIndex);
            int endRow = FindLastValidRow(lObjSheetAssignment, lPowerTableList, lIStartRowIndex);

            PowerMergeSheet sheet = new PowerMergeSheet();
            CreateTables(lObjSheetAssignment, lPowerTableList, lIStartRowIndex, endRow, sheet);

            _ = new PowerMerge();
            sheet.PowerMerge = ReadPowerMerge(sheet);

            return sheet;
        }

        private static bool FindFirstDataLocation(ExcelWorksheet excelWorksheet, out int lIStartRowIndex, out int lIStartColumnIndex)
        {
            lIStartRowIndex = 0;
            lIStartColumnIndex = 0;
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    if (EpplusExtensions.GetCellValue(excelWorksheet, i, j)?.Length != 0)
                    {
                        lIStartRowIndex = i;
                        lIStartColumnIndex = j;
                        break;
                    }
                }
                if (lIStartRowIndex != 0)
                {
                    break;
                }
            }
            return lIStartRowIndex != 0;
        }

        private static Hashtable GetPowerTableList(ExcelWorksheet excelWorksheet, int lIStartRowIndex, int lIStartColumnIndex)
        {
            Hashtable lPowerTableList = [];
            for (int i = lIStartColumnIndex; i <= excelWorksheet.Dimension.End.Column; i++)
            {
                string lStrValue = EpplusExtensions.GetCellValue(excelWorksheet, lIStartRowIndex, i);
                if (lStrValue.Length != 0 && !lPowerTableList.Contains(lStrValue))
                {
                    lPowerTableList.Add(lStrValue.ToUpper(), i);
                }
            }
            return lPowerTableList;
        }

        private static int FindLastValidRow(ExcelWorksheet excelWorksheet, Hashtable hashtable, int lIStartRowIndex)
        {
            var lEmpty = new Dictionary<string, Dictionary<int, bool>>();
            foreach (string powerTable in hashtable.Keys)
            {
                lEmpty.Add(powerTable, []);
                int lITableStartColumnIndex = int.Parse(hashtable[powerTable]!.ToString()!);
                for (int i = lIStartRowIndex + 2; i <= excelWorksheet.Dimension.End.Row; i++)
                {
                    string lPowerMerge1 = EpplusExtensions.GetCellValue(excelWorksheet, i, lITableStartColumnIndex);
                    string lPowerMerge2 = EpplusExtensions.GetCellValue(excelWorksheet, i, lITableStartColumnIndex + 1);
                    string lPowerMerge3 = EpplusExtensions.GetCellValue(excelWorksheet, i, lITableStartColumnIndex + 2);
                    lEmpty[powerTable].Add(i, lPowerMerge1?.Length == 0 && lPowerMerge2?.Length == 0 && lPowerMerge3?.Length == 0);
                }
            }

            int endRow = excelWorksheet.Dimension.End.Row;
            int blankCnt = 0;
            for (int i = lIStartRowIndex + 2; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                bool bothempty = true;
                foreach (string powerTable in hashtable.Keys)
                {
                    if (!lEmpty[powerTable][i])
                    {
                        bothempty = false;
                    }
                }
                blankCnt = bothempty ? blankCnt + 1 : 0;

                if (blankCnt >= 2)
                {
                    endRow = i - 2;
                    break;
                }
            }
            return endRow;
        }

        private static void CreateTables(ExcelWorksheet excelWorksheet, Hashtable hashtable, int lIStartRowIndex, int endRow, PowerMergeSheet powerMergeSheet)
        {
            foreach (string powerTable in hashtable.Keys)
            {
                int lITableStartColumnIndex = int.Parse(hashtable[powerTable]!.ToString()!);
                var lPowerMergeDataTable = new DataTable(powerTable);
                lPowerMergeDataTable.Columns.Add(PowerMergeSheet.ConHeaderNo);
                lPowerMergeDataTable.Columns.Add(PowerMergeSheet.ConHeaderNetName);
                lPowerMergeDataTable.Columns.Add(PowerMergeSheet.ConHeaderBallName);
                for (int i = lIStartRowIndex + 2; i <= endRow; i++)
                {
                    DataRow lPowerMergeDataRow = lPowerMergeDataTable.NewRow();
                    lPowerMergeDataRow[PowerMergeSheet.ConHeaderNo] = EpplusExtensions.GetCellValue(excelWorksheet, i, lITableStartColumnIndex);
                    lPowerMergeDataRow[PowerMergeSheet.ConHeaderNetName] = EpplusExtensions.GetCellValue(excelWorksheet, i, lITableStartColumnIndex + 1);
                    lPowerMergeDataRow[PowerMergeSheet.ConHeaderBallName] = EpplusExtensions.GetCellValue(excelWorksheet, i, lITableStartColumnIndex + 2);
                    lPowerMergeDataTable.Rows.Add(lPowerMergeDataRow);
                }
                powerMergeSheet.Tables.Add(lPowerMergeDataTable);
            }
        }

        public static PowerMerge ReadPowerMerge(PowerMergeSheet powerMergeSheet)
        {
            var powerMerge = new PowerMerge();
            if (powerMergeSheet == null)
            {
                return powerMerge;
            }

            DataTable? ftTable = powerMergeSheet.GetTable("FT");
            DataTable? cpTable = powerMergeSheet.GetTable("WS") ?? powerMergeSheet.GetTable("CP");
            CreateFtCpMapping(ftTable, cpTable, powerMerge);

            return powerMerge;
        }

        private static void CreateFtCpMapping(DataTable? ftTable, DataTable? cpTable, PowerMerge powerMerge)
        {
            if (ftTable == null || cpTable == null)
            {
                return;
            }

            string ftNetName = "";
            string cpNetName = "";
            for (int i = 0; i < ftTable.Rows.Count; i++)
            {
                string ftBallName = ftTable.Rows[i][2].ToString()!.ToUpper();
                UpdateFtCpNetNames(ftTable, cpTable, i, ftBallName, ref ftNetName, ref cpNetName);
                string cpBumpName = cpTable.Rows[i][2].ToString()!.ToUpper();
                AddFtCpMappings(powerMerge, ftBallName, cpBumpName, ftNetName, cpNetName);
            }
        }

        private static void UpdateFtCpNetNames(DataTable ftTable, DataTable cpTable, int i, string ftBallName, ref string ftNetName, ref string cpNetName)
        {
            string? ftNetVal = ftTable.Rows[i][1].ToString();
            if (!string.IsNullOrEmpty(ftNetVal))
            {
                ftNetName = ftNetVal.ToUpper();
            }
            else if (ftBallName == "VSS")
            {
                ftNetName = "";
            }

            string? cpNetVal = cpTable.Rows[i][1].ToString();
            if (!string.IsNullOrEmpty(cpNetVal))
            {
                cpNetName = cpNetVal.ToUpper();
            }
        }

        private static void AddFtCpMappings(PowerMerge powerMerge, string ftBallName, string cpBumpName, string ftNetName, string cpNetName)
        {
            if (cpNetName.Length != 0 && cpNetName != "N/A" && !powerMerge.FtCpMapping.ContainsKey(cpNetName))
            {
                powerMerge.FtCpMapping.Add(cpNetName, ftNetName);
            }

            if (ftBallName.Length != 0 && ftBallName != "N/A" && !powerMerge.FtPowers.ContainsKey(ftBallName))
            {
                powerMerge.FtPowers.Add(ftBallName, ftNetName);
            }

            if (cpBumpName.Length != 0 && cpBumpName != "N/A" && !powerMerge.CpPowers.ContainsKey(cpBumpName))
            {
                powerMerge.CpPowers.Add(cpBumpName, cpNetName);
            }

            if (cpBumpName.Length != 0 && cpBumpName != "N/A" && !powerMerge.CpBumpToFtNetMapping.ContainsKey(cpBumpName))
            {
                powerMerge.CpBumpToFtNetMapping.Add(cpBumpName, ftNetName);
            }

            if (ftBallName.Length != 0 && ftBallName != "N/A" && !powerMerge.FtBallToCpNetMapping.ContainsKey(ftBallName))
            {
                powerMerge.FtBallToCpNetMapping.Add(ftBallName, cpNetName);
            }
        }
    }

    public class PowerMerge
    {
        //FT: NetName<-->BallName
        public Dictionary<string, string> FtPowers = [];
        //WS: NetName<-->BumpName
        public Dictionary<string, string> CpPowers = [];
        // FT<-->WS power mapping
        public Dictionary<string, string> FtCpMapping = [];
        public Dictionary<string, string> CpBumpToFtNetMapping = [];
        public Dictionary<string, string> FtBallToCpNetMapping = [];

        public void GetCpFtNetName(string pinName, ref string cpNetName, ref string ftNetName)
        {
            // Get CP NetName by PowerName used in patInfo
            if (CpPowers.ContainsValue(pinName))
            {
                cpNetName = pinName;
            }
            else if (CpPowers.TryGetValue(pinName, out string? power))
            {
                cpNetName = power;
            }
            else if (FtBallToCpNetMapping.TryGetValue(pinName, out string? value))
            {
                cpNetName = value;
            }

            // Get FT NetName by PowerName used in patInfo
            if (FtPowers.ContainsValue(pinName))
            {
                ftNetName = pinName;
            }
            else if (FtPowers.TryGetValue(pinName, out string? power2))
            {
                ftNetName = power2;
            }
            else if (CpBumpToFtNetMapping.TryGetValue(pinName, out string? value2))
            {
                ftNetName = value2;
            }

            // Get CP/FT NetName by FtCp mapping when only one NetName can be found using CP mapping and FT mapping
            if (cpNetName.Length != 0 && !IsNCorNA(cpNetName) && ftNetName.Length == 0 && FtCpMapping.TryGetValue(cpNetName, out string? value1))
            {
                if (!IsNCorNA(value1))
                {
                    ftNetName = value1;
                }
            }
            if (ftNetName.Length != 0 && !IsNCorNA(ftNetName) && cpNetName.Length == 0 && FtCpMapping.ContainsValue(ftNetName))
            {
                string netName = ftNetName;
                List<string> allCpName = [.. FtCpMapping.Where(q => q.Value.EqualsIgnoreCase(netName)).Select(q => q.Key)];
                cpNetName = string.Join(",", allCpName);
            }
        }

        public static bool IsNCorNA(string pinName)
        {
            return pinName.EqualsIgnoreCase("NC") || pinName.EqualsIgnoreCase("N/A");
        }
    }
}
