using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Basic.Business.GenAc.AcGenerator.Business;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using IgxlLib.IgxlBase;

using TestPlanLib.HardIpDc.BaseData;
using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz
{
    public abstract class InsRowGenerator
    {
        protected HardIpSheet HardIpSheet { get; }
        private const string RegRetestSetting = @"retest:(?<setting>([^\&]+))";

        protected readonly HardIpInputData HardIpInputData;
        protected string SheetName = string.Empty;
        protected HardIpPattern Pattern;
        public string BlockName = string.Empty;
        public string SubBlockName = string.Empty;
        protected string TimingAc = string.Empty;
        protected string InstNameSubStr = string.Empty;
        protected bool NoPattern = false;
        protected Function VbtFunction;
        public HardIpCategoryDef HardIpDcSetting = null;
        protected string Lastpatname = string.Empty;
        protected string LastPatFailAction { get; set; } = string.Empty;
        public string LabelVoltage = string.Empty;
        public Dictionary<string, string> InitOriginalItems = new Dictionary<string, string>();
        public HardIpPattern Pat
        {
            set
            {
                Pattern = value;
                SetBasicInfoByPattern(value);
            }
            get
            {
                return Pattern;
            }
        }


        protected InsRowGenerator(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet, string sheetName)
        {
            HardIpSheet = hardIpSheet;
            HardIpInputData = hardIpInputData;
            SheetName = sheetName;
        }

        #region Main Methods
        /// <summary>
        /// Generate Instance rows for one pattern
        /// </summary>
        /// <returns></returns>
        public abstract List<InstanceRow> GenInsRows();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pattern"></param>
        protected abstract void SetBasicInfoByPattern(HardIpPattern pattern);

        /// <summary>
        /// Gen ReTest instance Row 
        /// </summary>
        /// <param name="instanceRow"></param>
        /// <returns></returns>
        internal InstanceRow GenReTestRow(InstanceRow instanceRow)
        {
            var retestInstRow = new InstanceRow();
            string setting = Regex.Match(Pattern.MiscInfo, RegRetestSetting, RegexOptions.IgnoreCase).Groups["setting"].ToString();
            retestInstRow.SheetName = SheetName;
            retestInstRow.TestName = instanceRow.TestName.Replace(LabelVoltage, HardIpConstData.PrefixReTest + LabelVoltage);
            retestInstRow.VbtName = instanceRow.VbtName;
            retestInstRow.ArgList = instanceRow.ArgList;
            retestInstRow.Args = instanceRow.Args;
            retestInstRow.VbtType = instanceRow.VbtType;
            retestInstRow.DcCategory = instanceRow.DcCategory;
            retestInstRow.DcSelector = instanceRow.DcSelector;
            retestInstRow.AcCategory = instanceRow.AcCategory;
            retestInstRow.AcSelector = instanceRow.AcSelector;
            retestInstRow.TimeSets = instanceRow.TimeSets;
            retestInstRow.PinLevels = instanceRow.PinLevels;
            retestInstRow.Overlay = instanceRow.Overlay;
            int index = retestInstRow.ArgList.Split(',').ToList().FindIndex(s => s.Equals(VbtFunctionParas.InterposePrePat, StringComparison.OrdinalIgnoreCase));
            if (index > -1)
            {
                string charInputString = retestInstRow.Args[index] + "," + setting;
                retestInstRow.Args[index] = charInputString.Trim(',');
            }
            return retestInstRow;
        }

        #endregion

        #region Generate each columns Methods

        /// <summary>
        /// Create Type, default: VBT
        /// </summary>
        /// <returns></returns>
        protected string CreateType()
        {
            return VbtFunction.Type;
        }

        /// <summary>
        /// Create Name: vbt function name
        /// </summary>
        /// <returns></returns>
        protected string CreateVbtName()
        {
            if (VbtFunction.Type.ToUpper() == ".NET")
            {
                return VbtFunction.FullFunctionName;
            }
            return VbtFunction.FunctionName;
        }

        /// <summary>
        /// Create ArgList: vbt function Parameters
        /// </summary>
        /// <returns></returns>
        protected string CreateArgList()
        {
            return VbtFunction.Parameters;
        }

        /// <summary>
        /// Create Args: vbt function ArgList
        /// </summary>
        /// <returns></returns>
        protected List<string> CreateArgs()
        {
            var setArgValueMain = new SetArgValueMain(HardIpInputData, HardIpSheet);
            if (VbtFunction.Parameters != "")
            {
                setArgValueMain.SetArgsValue(Pattern, ref VbtFunction, LabelVoltage);
            }
            return VbtFunction.ArgList.Select(p => p).ToList();
        }

        /// <summary>
        /// Create AcSelector
        /// if exist in PatternSummary, use what in summary
        /// else use default: Typ
        /// </summary>
        /// <returns></returns>
        protected string CreateAcSelector()
        {
            string specacselc = GetSpecifyInfo(Pattern.ForceCondition.ForceCondition, "ACSelector");
            if (!string.IsNullOrEmpty(specacselc))
            {
                return specacselc;
            }

            string patternName = Pattern.Pattern.GetLastPayload();
            if (VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)) || VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                patternName = Lastpatname;
            }


            if (!AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(patternName) &&
                !SearchInfo.IsHardipRtosSheet(Pattern.SheetName))
            {
                return "";
            }

            string acSelector = HardIpConstData.SelectTyp;

            if (!string.IsNullOrEmpty(Pattern.AcSelectorUsed))
            {
                string newAcSelector = GetAcSelector(LabelVoltage, Pattern.AcSelectorUsed);
                if (!string.IsNullOrEmpty(newAcSelector))
                {
                    acSelector = newAcSelector;
                }
            }
            return acSelector;
        }

        /// <summary>
        /// Create DcSelector
        /// HV:Max
        /// LV:Min
        /// NV:Typ
        /// </summary>
        /// <returns></returns>
        protected string CreateDcSelector()
        {
            string dcSelector;
            switch (LabelVoltage)
            {
                case HardIpConstData.LabelHv:
                    dcSelector = HardIpConstData.SelectMax;
                    break;
                case HardIpConstData.LabelLv:
                    dcSelector = HardIpConstData.SelectMin;
                    break;
                case HardIpConstData.LabelNv:
                    dcSelector = HardIpConstData.SelectTyp;
                    break;
                default:
                    dcSelector = HardIpConstData.SelectTyp;
                    break;
            }

            if (!string.IsNullOrEmpty(Pattern.DcSelectorUsed))
            {
                string newdcSelector = GetDcSelector(LabelVoltage, Pattern.DcSelectorUsed);
                if (!string.IsNullOrEmpty(newdcSelector))
                {
                    dcSelector = newdcSelector;
                }
            }
            return dcSelector;
        }
        #endregion

        /// <summary>
        /// Create AcCategory:
        /// if exist in PatternSummary, use what in summary
        /// if specified Timing AC, use [BlockName]_[TimingAc]
        /// else use default:HardIp
        /// </summary>
        /// <returns></returns>
        protected string CreateHardIpAcCategory(string timeSets)
        {
            string specaccate = GetSpecifyInfo(Pattern.ForceCondition.ForceCondition, "AC");
            if (!string.IsNullOrEmpty(specaccate))
            {
                return specaccate;
            }

            string patternName = Pattern.Pattern.GetLastPayload();
            if (VbtFunctionLibShared.EfusePrewriteFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase))
                || VbtFunctionLibShared.EfuseReadFunctionList.Exists(f => f.Equals(Pattern.FunctionName, StringComparison.CurrentCultureIgnoreCase)))
            {
                patternName = Lastpatname;
            }

            string ac = GetSpecifyInfo(Pattern.AcCategory, "AC");
            if (!string.IsNullOrEmpty(ac))
            {
                return ac;
            }

            if (Pattern.FunctionName.Equals(VbtFunctionLibShared.RtosBootUp, StringComparison.OrdinalIgnoreCase) ||
                Pattern.FunctionName.Equals(VbtFunctionLibShared.RtosRunScenarioT, StringComparison.OrdinalIgnoreCase))
            {
                return AcGenerator.SpiRom;
            }

            if (!AcTSetCategoryMapSingleton.Instance().PatternList.ContainsKey(patternName))
            {
                return "";
            }

            string timeSet2Cat = !string.IsNullOrEmpty(Pattern.TimeSetUsed.McgSetting) ?
                AcTSetCategoryMapSingleton.Instance().GetCategory(Pattern.TimeSetUsed.TimeSet, BlockType.HardIp) :
                AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets, BlockType.HardIp);
            if (timeSet2Cat == "TBD")
            {
                if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    int splitunderline = Pat.SheetName.IndexOf("_", StringComparison.Ordinal) + 1;
                    string blockName = splitunderline == 0 ? "" : Pat.SheetName.Substring(splitunderline, Pat.SheetName.Length - splitunderline);
                    string targetAc = NwireSingleton.Instance().SettingInfo.ReferenceFlow(NwireSingleton.Instance().SettingInfo.SettingTable).FirstOrDefault(p => p.ToUpper().Equals(blockName));
                    if (targetAc != null)
                    {
                        return targetAc;
                    }
                }

                timeSet2Cat = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSets);
            }
            if (!string.IsNullOrEmpty(TimingAc))
            {
                return BlockName + "_" + timeSet2Cat + "_" + TimingAc;
            }

            return timeSet2Cat;
        }

        internal string GetSpecifyInfo(string dcInfo, string key)
        {
            foreach (string info in dcInfo.Split(';'))
            {
                string[] dcArray = info.Split(':');
                if (dcArray.Length != 2)
                {
                    continue;
                }

                if (!dcArray[0].Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return dcArray[1];
            }
            return "";
        }

        internal string GetAcSelector(string labelVoltage, string acInfo)
        {
            string acSelector = "";
            foreach (string info in acInfo.Split(';'))
            {
                string[] acArray = info.Split(':');
                if (acArray.Length != 3)
                {
                    continue;
                }

                if (acArray[1].Equals(labelVoltage, StringComparison.OrdinalIgnoreCase))
                {
                    acSelector = acArray[2];
                }

                if (acArray[1].Equals(HardIpConstData.LabelAll, StringComparison.OrdinalIgnoreCase))
                {
                    acSelector = acArray[2];
                }
            }
            if (acSelector.Equals(HardIpConstData.SelectMax, StringComparison.OrdinalIgnoreCase))
            {
                acSelector = HardIpConstData.SelectMax;
            }

            if (acSelector.Equals(HardIpConstData.SelectMin, StringComparison.OrdinalIgnoreCase))
            {
                acSelector = HardIpConstData.SelectMin;
            }

            if (acSelector.Equals(HardIpConstData.SelectTyp, StringComparison.OrdinalIgnoreCase))
            {
                acSelector = HardIpConstData.SelectTyp;
            }

            return acSelector;
        }

        internal string GetDcSelector(string labelVoltage, string dcInfo)
        {
            string dcSelector = "";
            foreach (string info in dcInfo.Split(';'))
            {
                string[] dcArray = info.Split(':');
                if (dcArray.Length != 3)
                {
                    continue;
                }

                if (dcArray[1].Equals(labelVoltage, StringComparison.OrdinalIgnoreCase))
                {
                    dcSelector = dcArray[2];
                }

                if (dcArray[1].Equals(HardIpConstData.LabelAll, StringComparison.OrdinalIgnoreCase))
                {
                    dcSelector = dcArray[2];
                }
            }
            if (dcSelector.Equals(HardIpConstData.SelectMax, StringComparison.OrdinalIgnoreCase))
            {
                dcSelector = HardIpConstData.SelectMax;
            }

            if (dcSelector.Equals(HardIpConstData.SelectMin, StringComparison.OrdinalIgnoreCase))
            {
                dcSelector = HardIpConstData.SelectMin;
            }

            if (dcSelector.Equals(HardIpConstData.SelectTyp, StringComparison.OrdinalIgnoreCase))
            {
                dcSelector = HardIpConstData.SelectTyp;
            }

            return dcSelector;
        }
    }
}
