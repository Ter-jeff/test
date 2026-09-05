using System.Collections.Generic;

using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Static;

using TestPlanLib.Efuse.Input;

namespace Automation.InputManager.Data
{
    public class EFuseInputData : InputDataBase
    {
        public EfuseArraySizeSheet EfuseArraySizeSheet { get; set; }
        public EfuseReadSheet EfuseReadSheet { get; set; }
        public List<BitDefTable> EfuseBitDefTables { get; set; } = TestPlanStatic.BitDefTables;
        public List<BitDefBankRange> EfuseBitDefBankRange { get; set; } = new List<BitDefBankRange>();
        public List<EfuseConfigMainSheet> EfuseConfigMainSheets { get; set; } = new List<EfuseConfigMainSheet>();
        public int EfuseDatabaseRevision { get; set; } = -1;
        public List<EfusePatternRow> EfusePatternRows { get; set; } = new List<EfusePatternRow>();
        public ScghData EfuseScghSheet { get; set; }
        public EfuseBkmInfoReader EfuseBkmInfoTable { get; set; }
    }
}
