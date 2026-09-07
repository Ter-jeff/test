using System;
using System.Collections.Generic;
using System.Linq;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using NLog;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateBinTable
    {
        Logger Logger = LogManager.GetCurrentClassLogger();
        private BinTableRow _binShmooAlarm;

        public UpdateBinTable()
        {
            _binShmooAlarm = new BinTableRow();
            _binShmooAlarm.Name = "Bin_Central_Gating_Rule_By_Flow";
            _binShmooAlarm.ItemList = "F_Shmoo_Alarm";
            _binShmooAlarm.Op = "AND";
            _binShmooAlarm.Sort = "9988";
            _binShmooAlarm.Bin = "15";
            _binShmooAlarm.Result = "Fail-stop";
            _binShmooAlarm.Items = new List<string> { "T" };
        }
        //private BinTableRow _binShmooAlarm = new BinTableRow
        //{
        //    Name = "Bin_Central_Gating_Rule_By_Flow",
        //    ItemList = "F_Shmoo_Alarm",
        //    Op = "AND",
        //    Sort = "9988",
        //    Bin = "15",
        //    Result = "Fail-stop",
        //    Items = { "T" }
        //};
        public BinTableSheet Work(List<BinTableSheet> binTableSheets)
        {
            var binTableSheet = binTableSheets.FirstOrDefault(x => x.Name.Equals("Bin_Table", StringComparison.OrdinalIgnoreCase));
            if (binTableSheet == null)
            {
                return null;
            }

            AddOrInsert(binTableSheet, _binShmooAlarm);
            return binTableSheet;
        }
        private Boolean IsInsert(BinTableRow targetBin, string insertFlag)
        {
            var binItems = targetBin.ItemList.Split(',').Select(x => x.Trim().ToUpper());

            if (binItems.Any(x => x.Equals(insertFlag.ToUpper())))
                return false;
            else
                return true;
        }

        private void AddOrInsert(BinTableSheet binTable, BinTableRow bin)
        {
            var targetBin = binTable.Rows.FirstOrDefault(x => x.Name.Equals(bin.Name, StringComparison.OrdinalIgnoreCase) && x.IsBackup == false);
            if (targetBin != null)
            {
                foreach (var flag in bin.ItemList.Split(',').Select(x => x.Trim()))
                {
                    if (IsInsert(targetBin, flag))
                    {
                        targetBin.ItemList = targetBin.ItemList += "," + flag;
                        targetBin.Items.Add("T");
                    }
                }
            }
            else
            {
                binTable.AddRow(bin);
            }
        }
    }
}
