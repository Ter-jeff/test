using System.Collections.Generic;

using CommonReaderLib;

namespace TestPlanLib.Basic
{
    public class TimeSettingSheet : MySheet
    {
        public List<TimeSettingRow> Rows { set; get; }
        public int IndexSymbol = -1;

        #region Constructor
        public TimeSettingSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = [];
        }
        #endregion
    }

    public class TimeSettingRow : MyRow
    {
        public string Symbol { set; get; } = "";
        public Dictionary<string, string> SymbolValues = [];

        #region Constructor
        public TimeSettingRow(string sheetName = "")
        {
            SheetName = sheetName;
        }
        #endregion
    }
}
