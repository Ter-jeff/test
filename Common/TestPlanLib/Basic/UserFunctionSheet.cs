using System.Collections.Generic;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;

using TestPlanLib.VbtLib;

namespace TestPlanLib.Basic
{
    public class UserFunctionSheet
    {
        public List<UserFunctionTableRow> Rows = [];

        public void ArgumentSetting(string key, Function function)
        {
            if (Rows.Exists(x => x.UserFunction.EqualsIgnoreCase(key) && x.ArgumentSetting.Count != 0))
            {
                UserFunctionTableRow userFunctionRow = Rows.Find(x => x.UserFunction.EqualsIgnoreCase(key))!;
                foreach (KeyValuePair<string, string> arg in userFunctionRow.ArgumentSetting)
                {
                    function.SetParamValue(arg.Key, arg.Value);
                }
            }
        }

        public FlowRows CallAfterInstance(string key)
        {
            FlowRows flowRows = [];
            if (Rows.Exists(x => x.UserFunction.EqualsIgnoreCase(key) && !string.IsNullOrEmpty(x.CallAfterInstance)))
            {
                UserFunctionTableRow userFunctionRow = Rows.Find(x => x.UserFunction.EqualsIgnoreCase(key) && !string.IsNullOrEmpty(x.CallAfterInstance))!;

                string[] parts = userFunctionRow.CallAfterInstance.Split(':');
                if (parts.Length >= 2 &&
                    !string.IsNullOrEmpty(parts[0]) &&
                    !string.IsNullOrEmpty(parts[1]))
                {
                    flowRows.Add(new FlowRow { Opcode = parts[0], Parameter = parts[1] });
                }
            }
            return flowRows;
        }
    }

    public class UserFunctionTableRow
    {
        public string UserFunction = "";
        public Dictionary<string, string> ArgumentSetting = [];
        public string MultiFstpSetting = "";
        public string CallAfterInstance = "";
    }
}
