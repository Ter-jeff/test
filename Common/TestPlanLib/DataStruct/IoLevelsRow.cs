using System.Collections.Generic;

using CommonLib.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.Utility;

namespace TestPlanLib.DataStruct
{
    public class IoLevelsRow
    {
        #region Field
        public bool IsGroupPin;
        public bool IsTheSameRow;
        public string Domain = "";
        #endregion

        #region Properity
        public string SourceSheetName { set; get; } = "";
        public int RowNum { get; set; }
        public string Type { set; get; }
        public string PinName { set; get; }
        public string Fsdd { set; get; }
        public List<IoLevelsItem> IoLevelDate { set; get; }
        #endregion

        #region Constructor
        public IoLevelsRow()
        {
            Type = "";
            PinName = "";
            Fsdd = "";
            IoLevelDate = [];
        }

        public IoLevelsRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Type = "";
            PinName = "";
            Fsdd = "";
            IoLevelDate = [];
        }
        #endregion

        public static GlobalSpec GetGlobalSpec(string vdd, string value, string domain, string level, bool isTheSameDomain, string type = "")
        {
            string glbSymbol = GetGlobalSpecName(vdd, domain, level, isTheSameDomain, type);
            return new GlobalSpec(glbSymbol, SpecFormat.GenSpecValueSingleValue(value));
        }

        public static string GetGlobalSpecName(string vdd, string domain, string level, bool isTheSameDomain, string type = "")
        {
            domain = Combination.CombineByUnderLine(domain, type);
            string prefix = string.IsNullOrEmpty(vdd) ? "VIN_0v_" : "VIN_" + vdd.Replace(".", "p") + "v_";
            string glbSymbol = isTheSameDomain ? prefix + SpecFormat.GenGlbSpecSymbol(domain) : prefix + SpecFormat.GenGlbSpecSymbol(domain + "_" + level);
            return glbSymbol;
        }
    }
}
