using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using CommonReaderLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace TestPlanLib.Basic
{
    public class UfInstanceTable : MySheet
    {
        public List<UfInstanceRow> Rows = [];
        private List<string> InstanceList { get { return Rows.ConvertAll(x => x.TestName); } }

        public InstanceSheet CreateInstanceSheet()
        {
            var sheet = new InstanceSheet("TestInst_UF");
            foreach (UfInstanceRow row in Rows)
            {
                var tmpRow = new InstanceRow
                {
                    TestName = row.TestName,
                    DcCategory = row.DcSpec.Split(',').First(),
                    DcSelector = row.DcSpec.Split(',').Last(),
                    AcCategory = row.AcSpec.Split(',').First(),
                    AcSelector = row.AcSpec.Split(',').Last(),
                    PinLevels = row.PinLevels,
                    TimeSets = row.TimeSet,
                    VbtType = row.Type,
                    ArgList = row.ArgList,
                    VbtName = row.Module
                };
                var argList = new List<string>();
                foreach (string arg in row.Arg)
                {
                    argList.Add(arg);
                }
                tmpRow.Args = argList;
                sheet.AddRow(tmpRow);
            }
            return sheet;
        }

        public bool CheckExist(string instanceName)
        {
            return InstanceList.Exists(x => x.EqualsIgnoreCase(instanceName));
        }

        public bool CheckPowerUpExist()
        {
            return InstanceList.Exists(x => x.ContainsIgnoreCase("powerup") && !x.EqualsIgnoreCase("PowerUp_EVS"));
        }

        public List<int> GetPowerUpInstanceRowNum()
        {
            var rowNumList = new List<int>();
            for (int i = 0; i < Rows.Count; i++)
            {
                UfInstanceRow row = Rows[i];
                if (row.TestName.ContainsIgnoreCase("powerup") && !row.TestName.EqualsIgnoreCase("PowerUp_EVS"))
                {
                    rowNumList.Add(i + 2);
                }
            }
            return rowNumList;
        }
    }

    public class UfInstanceRow
    {
        public string TestName = "";
        public string DcSpec = string.Empty;
        public string AcSpec = string.Empty;
        public string PinLevels = "";
        public string TimeSet = "";
        public string Module = string.Empty;
        public string ArgList = "";
        public List<string> Arg = [];
        public string Type
        {
            get { return Module.Contains('.') ? ".NET" : "VBT"; }
        }
    }
}
