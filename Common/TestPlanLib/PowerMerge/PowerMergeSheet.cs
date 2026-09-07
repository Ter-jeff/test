using System;
using System.Collections.Generic;
using System.Data;

using CommonLib.Extension;

namespace TestPlanLib.PowerMerge
{
    public class PowerMergeSheet : DataSet
    {
        public const string ConHeaderNo = "No";
        public const string ConHeaderNetName = "NET_NAME";
        public const string ConHeaderBallName = "Ball_NAME";
        public PowerMerge PowerMerge { get; set; } = new PowerMerge();

        #region Member Function

        public DataTable? GetTable(string tableName)
        {
            return Tables.Contains(tableName.ToUpper()) ? Tables[tableName.ToUpper()] : null;
        }

        public List<string> GetMergePins()
        {
            List<string> lPinList = [];
            foreach (DataTable dataTbl in Tables)
            {
                foreach (DataRow dataRow in dataTbl.Rows)
                {
                    if (!lPinList.Contains(dataRow[ConHeaderNetName].ToString()!))
                    {
                        lPinList.Add(dataRow[ConHeaderNetName].ToString()!);
                    }
                }
            }
            return lPinList;
        }

        public List<string> GetNetPinsByTable(string tableName)
        {
            List<string> lPinList = [];
            DataTable? table = GetTable(tableName);
            if (table != null)
            {
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!lPinList.Contains(dataRow[ConHeaderNetName].ToString()!))
                    {
                        lPinList.Add(dataRow[ConHeaderNetName].ToString()!);
                    }
                }
            }
            return lPinList;
        }

        public List<string> GetBallPinsByTable(string tableName)
        {
            List<string> lPinList = [];
            DataTable? table = GetTable(tableName);
            if (table != null)
            {
                foreach (DataRow dataRow in table.Rows)
                {
                    if (!lPinList.Contains(dataRow[ConHeaderBallName].ToString()!))
                    {
                        lPinList.Add(dataRow[ConHeaderBallName].ToString()!);
                    }
                }
            }
            return lPinList;
        }

        public string CovertNetNameFromBallName(string ballName, string job)
        {
            string netName = string.Empty;
            string currentNetName = string.Empty;
            DataTable table;
            if (Tables.Contains(job.ToUpper()))
            {
                table = Tables[job.ToUpper()]!;
                foreach (DataRow dataRow in table.Rows)
                {
                    string? netNameCell = dataRow[ConHeaderNetName].ToString();
                    if (!string.IsNullOrEmpty(netNameCell))
                    {
                        currentNetName = netNameCell;
                    }

                    if (dataRow[ConHeaderBallName].ToString()!.EqualsIgnoreCase(ballName))
                    {
                        netName = currentNetName;
                        break;
                    }
                }
            }
            else
            {
                throw new Exception($"Can not find Job : {job} in PowerMerge");
            }

            return netName;
        }

        public static DataTable CovertPowMergeFromPowerPowerTable(List<string> inputList)
        {
            int cnt = 1;
            DataTable table = new DataTable();
            table.Columns.Add(ConHeaderNo);
            table.Columns.Add(ConHeaderNetName);
            table.Columns.Add(ConHeaderBallName);
            foreach (string power in inputList)
            {
                DataRow row = table.NewRow();
                row[ConHeaderNo] = cnt++;
                row[ConHeaderNetName] = power;
                row[ConHeaderBallName] = power;
                table.Rows.Add(row);
            }
            return table;
        }

        #endregion
    }
}
