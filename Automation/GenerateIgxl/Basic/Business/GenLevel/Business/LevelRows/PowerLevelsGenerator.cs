using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenLevel.BassData;
using Automation.Singleton;
using Automation.Static;
using Automation.Utility.Basic;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using TestPlanLib.DataStruct;
using TestPlanLib.Static;

namespace Automation.GenerateIgxl.Basic.Business.GenLevel.Business.LevelRows
{
    internal class PowerLevelsGenerator : LevelsGenerator
    {
        private readonly MultiTestSettingSheetsSingleton _testSettings;
        private readonly PowerInfoSheet _powerInfo;
        private readonly PinMapSheet _pinMapSheet;
        private readonly bool _generateVSlewRate;
        private readonly bool _useSaveVoltage;
        private readonly string _chiplet;

        public PowerLevelsGenerator(MultiTestSettingSheetsSingleton testSettingSheetsSingleton, PowerInfoSheet powerInfo, PinMapSheet pinMapSheet, bool generateVSlewRate, bool useSaveVoltage, string chiplet = "")
        {
            _testSettings = testSettingSheetsSingleton;
            _powerInfo = powerInfo;
            _pinMapSheet = pinMapSheet;
            _generateVSlewRate = generateVSlewRate;
            _useSaveVoltage = useSaveVoltage;
            _chiplet = chiplet;
        }

        public override void UpdateLevelSheet(ref LevelSheet levelSheet, LevelData levelData)
        {
            CreatePowerLevel(levelSheet, levelData.SplitDcBlockSyntax);
        }

        private void CreatePowerLevel(LevelSheet levelSheet, string splitDcBlockSyntax)
        {
            bool conti = Regex.IsMatch(levelSheet.Name, "Conti|WalkingZ", RegexOptions.IgnoreCase);
            Dictionary<string, string> iFlodLimits = conti ? ContiPowerPinLimits() : null;

            if (_pinMapSheet.IsGroupExist(DcContiConst.PinGroupAllDcvi))
            {
                _pinMapSheet.GetGroup(DcContiConst.PinGroupAllDcvi).PinList.Select(x => x.PinName).ToList();
            }

            foreach (string powerPin in _testSettings.PowerPinList)
            {
                string pinName = powerPin;
                if (!_pinMapSheet.IsGroupExist(pinName) && !_pinMapSheet.IsPinExist(pinName))
                {
                    continue;
                }

                string comment = GetPowerInfoCheckingForLevelComment(pinName, _powerInfo);
                string contiIfold = "0.2";
                if (iFlodLimits != null && iFlodLimits.ContainsKey(pinName.ToUpper()))
                {
                    contiIfold = iFlodLimits[pinName.ToUpper()];
                }

                string vmain = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, "", "", splitDcBlockSyntax, _chiplet);  //=_VDD_CPU_VAR_C
                string valt = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, SpecFormat.DcSpecVopSuffix, "", splitDcBlockSyntax, _chiplet);  //=_VDD_CPU_VOP_VAR_C
                if (!_testSettings.VrsPowerPinList.Exists(s => s.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                {
                    valt = vmain;
                }

                string ifoldlevel = "";
                if (levelSheet.Name.ContainsIgnoreCase("EVS"))
                {
                    ifoldlevel = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, _powerInfo.ExistEvs ? "IfoldEVS" : "Ifold");  //=_VDD_CPU_Ifold_GLB
                }
                else if (levelSheet.Name.ContainsIgnoreCase("HardIP") || levelSheet.Name.ContainsIgnoreCase("IDS"))
                {
                    ifoldlevel = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, _powerInfo.ExistHip ? "IfoldHIP" : "Ifold");  //=_VDD_CPU_IfoldHIP_GLB
                }
                else if (conti)
                {
                    ifoldlevel = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, _powerInfo.ExistConti ? "IfoldConti" : "Ifold");  //=_VDD_CPU_Ifold_GLB
                }
                else
                {
                    ifoldlevel = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Ifold");  //=_VDD_CPU_Ifold_GLB
                }

                if (conti && GetIfoldFromPowerInfo(pinName) != "")
                {
                    if (Convert.ToDouble(contiIfold) > Convert.ToDouble(GetIfoldFromPowerInfo(pinName)))
                    {
                        string errorMessage =
                            $"Ifold level {ifoldlevel} defined in PowerInfo is greater than Ifold limit for its instrument, please check!!!";
                        ErrorReportManager.AddError(BasicErrorType.E_ContiIfold_01, levelSheet.Name, 1, 1, $"Ifold level {ifoldlevel} defined in PowerInfo is greater than Ifold limit for its instrument, please check!!!", new string[] { ifoldlevel });
                    }
                }
                AddPowerPins(levelSheet, pinName, vmain, valt, ifoldlevel, _generateVSlewRate, comment: comment);
            }
        }

        private Dictionary<string, string> ContiPowerPinLimits()
        {
            var contiPowerLimits = new Dictionary<string, string>();
            Dictionary<string, ChannelMapSheet> channelMapDic = TestProgram.IgxlWorkBk.ChannelMapSheets;
            foreach (ChannelMapSheet channelMap in channelMapDic.Values)
            {
                Dictionary<string, string> lconti = GetContiPowerPinLimits(channelMap);
                foreach (string pinname in lconti.Keys)
                {
                    if (contiPowerLimits.ContainsKey(pinname))
                    {
                        double first = Convert.ToDouble(contiPowerLimits[pinname]);
                        double second = Convert.ToDouble(lconti[pinname]);
                        if (first > second) // took the smaller one
                        {
                            contiPowerLimits[pinname] = second.ToString();
                        }
                    }
                    else
                    {
                        contiPowerLimits.Add(pinname, lconti[pinname]);
                    }
                }
            }
            return contiPowerLimits;
        }

        private Dictionary<string, string> GetContiPowerPinLimits(ChannelMapSheet channelMap)
        {
            var contiPowerLimits = new Dictionary<string, string>();
            foreach (ChannelMapRow chRow in channelMap.Rows)
            {
                if (contiPowerLimits.ContainsKey(chRow.DeviceUnderTestPinName.ToUpper()))
                {
                    continue;
                }

                string mergeN = Regex.Match(chRow.Type, @"Merged(?<N>\d+)").Groups["N"].ToString();
                double n = 1;
                if (mergeN != "")
                {
                    n = Convert.ToDouble(mergeN);
                }


                double iFoldLimit;
                if (Regex.IsMatch(chRow.Type, "DCVI", RegexOptions.IgnoreCase))
                {
                    iFoldLimit = 0.2 * n;//UVI80: 0.2A
                    contiPowerLimits.Add(chRow.DeviceUnderTestPinName.ToUpper(), iFoldLimit.ToString());
                }
                else if (Regex.IsMatch(chRow.Type, "DCVS", RegexOptions.IgnoreCase))
                {
                    iFoldLimit = chRow.InstrumentType.Equals("HexVS", StringComparison.CurrentCultureIgnoreCase) ? 0.3 : 0.2;  //HexVs: 0.3A, UVS256: 0.2A
                    iFoldLimit *= n;

                    contiPowerLimits.Add(chRow.DeviceUnderTestPinName.ToUpper(), iFoldLimit.ToString());
                }
            }
            return contiPowerLimits;
        }

        private protected string GetPowerInfoCheckingForLevelComment(string pinName, PowerInfoSheet powerInfo)
        {
            string result = "";
            if (powerInfo != null)
            {
                var powerInfoPin = powerInfo.Rows.Select(x => x.PinName.ToUpper()).ToList();
                if (!powerInfoPin.Contains(pinName.ToUpper()))
                {
                    result = $"Missing pin in PowerInfo for iFold: {pinName.ToUpper()}";
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_18, NeededSheets.PowerInfo, 0, 0, $"Missing pin in PowerInfo for iFold: {pinName.ToUpper()}", new string[] { pinName.ToUpper() }, new ErrorInfo() { Comments = new List<string>() { pinName } });
                }
            }
            return result;
        }

        private string GetIfoldFromPowerInfo(string pinName)
        {
            PowerInfoRow pinRow = _powerInfo.Rows.Find(s => s.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
            if (pinRow != null && double.TryParse(pinRow.Ifold.Trim(), out double _))
            {
                return pinRow.Ifold.Trim();
            }

            return "";
        }

        private void AddPowerPins(LevelSheet levelSheet, string pinName, string vmain, string valt, string ifoldlevel, bool generateVSlewRate = false, string comment = "")
        {
            bool isMultiTypePin = false;
            List<string> allPinType = PinService.GetPinTypeFromChannelMaps(pinName, _testSettings.PowerPinList, TestProgram.IgxlWorkBk.ChannelMapSheets);
            if (allPinType.Count == 0)
            {
                allPinType.Add("");
            }
            else
            {
                IEnumerable<string> dcvspins = allPinType.Where(p => Regex.IsMatch(p, "DCVS", RegexOptions.IgnoreCase));
                IEnumerable<string> dcvipins = allPinType.Where(p => Regex.IsMatch(p, "DCVI", RegexOptions.IgnoreCase));
                if (dcvspins.Any() && dcvipins.Any() && _powerInfo.Rows.Exists(p => p.PinName.Equals(pinName)))
                {
                    if (allPinType.Count >= 2)
                    {
                        allPinType.Clear();
                        allPinType.Add("DCVI");
                        allPinType.Add("DCVS");
                    }
                    isMultiTypePin = true;
                }
                else
                {
                    allPinType.RemoveRange(1, allPinType.Count - 1);
                }
            }

            foreach (string pinType in allPinType)
            {
                string newPinName = pinName;
                string newIfoldlevel = ifoldlevel;
                if (isMultiTypePin)
                {
                    newPinName += "_" + pinType;
                    newIfoldlevel = ifoldlevel.Replace(pinName, newPinName);
                }

                if (pinType.Equals("DCVI", StringComparison.OrdinalIgnoreCase) ||
                    pinType.Equals("DCVIMerged", StringComparison.OrdinalIgnoreCase))
                {
                    DcviPowerLevel lObjDcviPowerLevel = new DcviPowerLevel(newPinName, vmain, newIfoldlevel, "0", comment: comment);
                    levelSheet.AddDcviPowerPinLevel(lObjDcviPowerLevel);
                }
                else
                {
                    //For Levels_Mbist,change the value of vmain and valt
                    if (levelSheet.Name.Equals("Levels_" + LevelInitial.Mbist) && _useSaveVoltage)
                    {
                        (vmain, valt) = (valt, vmain);
                    }

                    if (generateVSlewRate && !Regex.IsMatch(levelSheet.Name, "EVS", RegexOptions.IgnoreCase))
                    {
                        PowerLevel lObjPowerLevel = new PowerLevel(newPinName, vmain, valt, newIfoldlevel, "0", "50", comment: comment);
                        levelSheet.AddPowerPinVSlewLevel(lObjPowerLevel);
                    }
                    else
                    {
                        PowerLevel lObjPowerLevel = new PowerLevel(newPinName, vmain, valt, newIfoldlevel, "0", comment: comment);
                        levelSheet.AddPowerPinLevel(lObjPowerLevel);
                    }
                }
            }
        }
    }
}
