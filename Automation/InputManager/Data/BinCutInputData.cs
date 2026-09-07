using System;
using System.Collections.Generic;
using System.Linq;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.PreCheck.AllParaData;
using Automation.Reader.ConfigFile.NamingRule.Base;

using CommonReaderLib.PatternListCsv;

using IgxlLib.IgxlBase;

using TestPlanLib.BinCut;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.Flow;
using TestPlanLib.BinCut.FlowNew;
using TestPlanLib.HardIpDc.BaseData;

namespace Automation.InputManager.Data
{
    public class BinCutInputData : InputDataBase
    {
        public BinCutFlowTables BinCutFlowTables { get; set; } = new BinCutFlowTables();
        public BinCutFlowTables BinCutFlowShadowTables { get; set; } = new BinCutFlowTables();
        public NewBinCutFlowTables NewBinCutFlowTables { get; set; } = new NewBinCutFlowTables();
        public BinCutFlowSheets BinCutPostFlowSheets { get; set; } = new BinCutFlowSheets();
        public NewBinCutFlowTables NewPostBinCutFlowTables { get; set; } = new NewBinCutFlowTables();
        public BinningTables BinningTables { get; set; } = new BinningTables();
        public IdsDistributionTable IdsdistributionTable { get; set; } = new IdsDistributionTable();
        public BinCutOrderSheet BinCutOrderSheet { get; set; }
        public List<BinCutInstanceSheet> BinCutInstanceSheets { get; set; } = new List<BinCutInstanceSheet>();
        public List<BinCutInstanceSheet> BinCutInstancePostSheets { get; set; } = new List<BinCutInstanceSheet>();
        public Dictionary<string, HardIpSheet> HardIpPatterns { get; set; } = new Dictionary<string, HardIpSheet>();
        public BinCutInstanceNamingSheet BinCutInstanceNamingSheet { get; set; }
        public ScanConfig Config { get; set; }

        public List<InstanceRow> GenHardipInstanceByPattern(string pattern, HardIpDcSheet hardIpDcSheet)
        {
            var newDic = new Dictionary<string, HardIpSheet>();
            foreach (KeyValuePair<string, HardIpSheet> dic in HardIpPatterns)
            {
                foreach (HardIpPattern row in dic.Value.Rows)
                {
                    if (row.Pattern.PatternSetList.SelectMany(x => x).Any(x => x.Equals(pattern, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        newDic.Add(dic.Key, new HardIpSheet { SheetName = dic.Key, Rows = new List<HardIpPattern> { row } });
                        break;
                    }
                }
            }
            var hardIpInputData = new HardIpInputData(new HardIpParaData(EnumBlock.HardIp))
            {
                HardIpDcSheet = hardIpDcSheet
            };
            var instanceGenerator = new InstanceGenerator(hardIpInputData);
            return instanceGenerator.GenInst(newDic).SelectMany(x => x.Rows).ToList();
        }
    }
}
