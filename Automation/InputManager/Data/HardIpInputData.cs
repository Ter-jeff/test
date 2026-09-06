using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.PreCheck.AllParaData;
using Automation.Utility.TpUpdate.HardIPEnableWordsUpdate;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;
using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.PowerMerge;

namespace Automation.InputManager.Data
{
    public class HardIpInputData : InputDataBase
    {
        public HardIpParaData HardIpParaData { get; }
        public Dictionary<string, string> VbtNameMapping { get; set; } = new Dictionary<string, string>();
        public HardIpDcSheet HardIpDcSheet { get; set; } = new HardIpDcSheet();
        public Dictionary<string, EnableWordTable> TtrTables { get; set; } = new Dictionary<string, EnableWordTable>();
        public Dictionary<string, string> EfuseBitDefBw { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, HardIpSheet> PlanDic { get; set; } = new Dictionary<string, HardIpSheet>(StringComparer.OrdinalIgnoreCase);
        public IdsMappingSheet IdsMappingSheet { get; internal set; }
        public PowerInfoSheet PowerInfoSheet { get; set; }
        public PowerMergeSheet PowerMergeSheet { get; set; }

        public List<HardIpRegAssign> HardIpRegAssigns { get; set; } = new List<HardIpRegAssign>();
        public ConcurrentBag<InterposeAssign> InterposeAssigns { get; set; } = new ConcurrentBag<InterposeAssign>();
        public ConfigData ConfigData { get; set; } = new ConfigData();

        public HardIpInputData(HardIpParaData hardIpParaData)
        {
            HardIpParaData = hardIpParaData;
        }
    }
}
