using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.SpecialSetting;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.HardIP;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow
{
    public abstract class FlowSheetGeneratorBase
    {
        protected HardIpInputData HardIpInputData { get; }
        private const string PrefixFlowSheet = "Flow_";
        private const string SuffixShmooFlowSheet = "_CZ2";
        private const string SuffixVtShmooFlowSheet = "_VT";
        protected string SheetName = string.Empty;
        public string FlowSheetName = string.Empty;
        //Pattern List from test plan 
        protected List<HardIpPattern> PatternList;
        //Pattern List after divided
        protected List<HardIpPattern> ExtendedPatList;
        protected FlowRowGeneratorBase FlowRowGenerator = null;

        protected FlowSheetGeneratorBase(HardIpInputData hardIpInputData, string sheetName, List<HardIpPattern> patternList)
        {
            HardIpInputData = hardIpInputData;
            SheetName = CommonGenerator.GetHardipSheetName(sheetName).ToUpper();
            PatternList = patternList;
        }

        public List<SubFlowSheet> GenerateFlowSheetForRfMain(List<FlowRow> subFlowRows)
        {
            var flowSheets = new List<SubFlowSheet>();
            FlowSheetName = PrefixFlowSheet + SheetName;
            var flowSheet = new SubFlowSheet(FlowSheetName, SheetName);
            flowSheet.AddRows(GenStartRows());
            flowSheet.AddRows(subFlowRows);
            flowSheet.AddRows(GenEndRows());
            flowSheets.Add(flowSheet);
            return flowSheets;
        }

        public List<SubFlowSheet> GenerateFlowSheet()
        {
            var flowSheets = new List<SubFlowSheet>();
            FlowSheetName = PrefixFlowSheet + SheetName;
            var flowSheet = new SubFlowSheet(FlowSheetName, SheetName);
            ExtendedPatList = DividePatterns();

            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                flowSheet.AddRows(GenNwireRows());
            }

            if (SheetName.StartsWith(HardIpConstData.PrefixWireless, StringComparison.OrdinalIgnoreCase))
            {
                flowSheet.AddRow(FlowRowGenerator.GenBlockDatalogFlow("Enable"));
            }


            flowSheet.AddRows(GenStartRows());
            flowSheet.AddRows(GenFlowBodyRows());

            if (SheetName.StartsWith(HardIpConstData.PrefixWireless, StringComparison.OrdinalIgnoreCase))
            {
                flowSheet.AddRow(FlowRowGenerator.GenBlockDatalogFlow("Disable"));
            }

            flowSheet.AddRows(GenEndRows());
            flowSheets.Add(flowSheet);
            return flowSheets;
        }

        public List<SubFlowSheet> GenerateFlowSheetForSplitCz()
        {
            List<string> typeForCz = LocalSpecs.Options.CharacterizationType.Split(',').ToList();
            var flowSheets = new List<SubFlowSheet>();
            var czExtendedPatList = new List<HardIpPattern>();
            var nonczExtendedPatList = new List<HardIpPattern>();
            ExtendedPatList = DividePatterns();


            foreach (HardIpPattern extended in ExtendedPatList)
            {
                if (extended.Pattern.IsMultiTimeDomain() &&
                    !(extended.MiscInfoDict.ContainsKey("Ref_SubBlock") && extended.MiscInfoDict["Ref_SubBlock"].Contains("#")))
                {
                    continue;
                }

                if (typeForCz.Exists(x => extended.Pattern.GetLastPayload().StartsWith(x, StringComparison.CurrentCultureIgnoreCase)) ||
                    extended.Pattern.InstancePatternName.All(name => typeForCz.Any(prefix => name.StartsWith(prefix, StringComparison.CurrentCultureIgnoreCase))) ||
                    HardIpService.IsCz2Only(extended.MiscInfo))
                {
                    czExtendedPatList.Add(extended);
                }
                else
                {
                    nonczExtendedPatList.Add(extended);
                }
            }

            if (czExtendedPatList.Count > 0)
            {
                foreach (HardIpPattern pattern in czExtendedPatList)
                {
                    if (typeForCz.Exists(x => pattern.Pattern.GetLastPayload().StartsWith(x, StringComparison.CurrentCultureIgnoreCase)) ||
                        pattern.Pattern.InstancePatternName.All(name => typeForCz.Any(prefix => name.StartsWith(prefix, StringComparison.CurrentCultureIgnoreCase))) ||
                        HardIpService.IsCz2Only(pattern.MiscInfo))
                    {
                        pattern.ForceCondition.IsShmooInCharFlow = true;
                        pattern.ForceCondition.IsShmooInProdInst = true;
                    }
                }

                FlowSheetName = PrefixFlowSheet + SheetName + "_CZ";
                var czFlowSheet = new SubFlowSheet(FlowSheetName, SheetName + "_CZ");
                ExtendedPatList = czExtendedPatList;
                czFlowSheet.AddRows(GenStartRows(FlowSheetName));
                czFlowSheet.AddRows(GenFlowBodyRows(true));
                czFlowSheet.AddRows(GenEndRows(FlowSheetName));
                flowSheets.Add(czFlowSheet);
            }

            ExtendedPatList = nonczExtendedPatList;
            FlowSheetName = PrefixFlowSheet + SheetName;
            var flowSheet = new SubFlowSheet(FlowSheetName, SheetName);
            if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                flowSheet.AddRows(GenNwireRows());
            }
            flowSheet.AddRows(GenStartRows());
            flowSheet.AddRows(GenFlowBodyRows());
            flowSheet.AddRows(GenEndRows());
            flowSheets.Add(flowSheet);

            return flowSheets;
        }

        public virtual SubFlowSheet GenerateShmooFlowSheet()
        {
            FlowSheetName = PrefixFlowSheet + SheetName + SuffixShmooFlowSheet;
            List<FlowRow> shmooFlowBodyRows = GenFlowBodyRows(true);
            if (shmooFlowBodyRows.Any())
            {
                var flowSheet = new SubFlowSheet(FlowSheetName, SheetName + SuffixShmooFlowSheet);
                flowSheet.AddRows(GenStartRows());
                flowSheet.AddRows(shmooFlowBodyRows);
                flowSheet.AddRows(GenShmooEndRows());
                return flowSheet;
            }
            return null;
        }

        public virtual SubFlowSheet GenerateVtShmooFlowSheet()
        {
            FlowSheetName = PrefixFlowSheet + SheetName + SuffixVtShmooFlowSheet;
            List<FlowRow> vtShmooFlowBodyRows = GenFlowBodyRows(false, true);
            if (vtShmooFlowBodyRows.Any())
            {
                var flowSheet = new SubFlowSheet(FlowSheetName, SheetName + SuffixVtShmooFlowSheet);
                flowSheet.AddRows(GenStartRows());
                flowSheet.AddRows(vtShmooFlowBodyRows);
                flowSheet.AddRows(GenShmooEndRows());
                return flowSheet;
            }
            return null;
        }

        public SubFlowSheet GenerateCz2MainFlow(List<SubFlowSheet> shmooflowSheets, string cz2SheetName)
        {
            FlowSheetName = cz2SheetName;
            var flowSheet = new SubFlowSheet(FlowSheetName, "HARDIP_CHAR");
            flowSheet.AddRows(GenStartRows(FlowSheetName));
            flowSheet.AddRows(GenCallRows(shmooflowSheets));
            flowSheet.AddRows(GenShmooEndRows());
            return flowSheet;
        }

        #region Abstract Methods
        protected abstract List<HardIpPattern> DividePatterns();

        protected abstract List<FlowRow> GenFlowBodyRows(bool shmooflag = false, bool vtShmooFlag = false);
        #endregion

        protected virtual List<FlowRow> GenNwireRows()
        {
            var flowRows = new List<FlowRow>();
            //Nwire rows
            FlowRow nwireRow = NwireSingleton.Instance().SettingInfo.GetNwireCall(FlowSheetName);
            if (!string.IsNullOrEmpty(nwireRow.Parameter))
            {
                flowRows.Add(nwireRow);
            }

            return flowRows;
        }

        #region Virtual Methods
        internal virtual List<FlowRow> GenStartRows(string sheetName = "")
        {
            var flowRows = new List<FlowRow>();
            InstanceSheet instSheet = TestProgram.IgxlWorkBk.InsSheets.Values.ToList().Find(p =>
                p.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));

            InstanceRow enableInst = instSheet?.Rows?.FirstOrDefault(x => x.TestName.Equals(FlowRowGenerator.GenHardIpEnableDisableName("Enable"), StringComparison.OrdinalIgnoreCase));

            if ((string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && enableInst != null) ||
                LocalSpecs.Options.Device == EnumDevice.RF)
            {
                if (Regex.IsMatch(FlowSheetName, HardIpConstData.PrefixHardIp + "|" + HardIpConstData.GpioBlockName + "|" + HardIpConstData.TmpsBlockName, RegexOptions.IgnoreCase) &&
                    !Regex.IsMatch(FlowSheetName, "init|nWire", RegexOptions.IgnoreCase))
                {
                    flowRows.Add(FlowRowGenerator.GenHardIpDatalogFlow("Enable"));
                }
            }

            //start row
            flowRows.Add(FlowRowGenerator.GenPrintStartRow(sheetName));
            return flowRows;
        }

        internal virtual List<FlowRow> GenEndRows(string sheetName = "")
        {
            var flowRows = new List<FlowRow>();

            InstanceSheet instSheet = TestProgram.IgxlWorkBk.InsSheets.Values.ToList().Find(p =>
                p.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            InstanceRow enableInst = instSheet?.Rows?.FirstOrDefault(x => x.TestName.Equals(FlowRowGenerator.GenHardIpEnableDisableName("Disable"), StringComparison.OrdinalIgnoreCase));

            if ((string.IsNullOrEmpty(LocalSpecs.CsLibraryFolder) && enableInst != null) ||
                LocalSpecs.Options.Device == EnumDevice.RF)
            {
                if (Regex.IsMatch(FlowSheetName, HardIpConstData.PrefixHardIp + "|" + HardIpConstData.GpioBlockName + "|" + HardIpConstData.TmpsBlockName, RegexOptions.IgnoreCase) &&
                    !Regex.IsMatch(FlowSheetName, "init|nWire", RegexOptions.IgnoreCase))
                {
                    flowRows.Add(FlowRowGenerator.GenHardIpDatalogFlow("Disable"));
                }
            }

            flowRows.Add(FlowRowGenerator.GenPrintStopRow(sheetName));
            flowRows.Add(FlowRowGeneratorBase.GenReturnRow());
            return flowRows;
        }

        protected virtual List<FlowRow> GenShmooEndRows()
        {
            var flowRows = new List<FlowRow>
            {
                FlowRowGenerator.GenPrintStopRow(), FlowRowGeneratorBase.GenReturnRow()
            };
            return flowRows;
        }

        protected virtual List<FlowRow> GenCallRows(List<SubFlowSheet> shmooflowSheets)
        {
            var flowRows = new List<FlowRow>();
            foreach (SubFlowSheet sheet in shmooflowSheets)
            {
                flowRows.Add(FlowRowGeneratorBase.GenCallRow(sheet.Name));
            }

            return flowRows;
        }
        #endregion

        #region Common Methods
        protected List<FlowRow> GenResetRelayRows(string labelVoltage = "")
        {
            var flowRows = new List<FlowRow>();
            var lastSetting = new Dictionary<string, string>();
            HardIpPattern lastPlanItem = ExtendedPatList.LastOrDefault();
            if (lastPlanItem != null)
            {
                lastSetting = lastPlanItem.NewRelaySetting;
            }

            if (lastSetting.Count > 0)
            {
                string lastEnable = !string.IsNullOrEmpty(labelVoltage) &&
                                !Regex.IsMatch(ExtendedPatList[ExtendedPatList.Count - 1].MiscInfo, HardIpConstData.RemoveNv,
                                    RegexOptions.IgnoreCase)
                ? HardIpConstData.PrefixHardIp + labelVoltage
                : "";
                foreach (KeyValuePair<string, string> item in lastSetting)
                {
                    string setting = item.Value;
                    flowRows.AddRange(RelaySettingMain.GenRelaySettingInJob(null, SearchInfo.ReverseRelaySetting(setting), item.Key, lastEnable));
                }
            }
            return flowRows;
        }
        #endregion
    }
}
