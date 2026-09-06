using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Base;
using Automation.Reader;
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

namespace Automation.GenerateIgxl.Basic.Business.GenGlobalSpec
{
    public class GlobalSpecGenerator : GeneratorBase
    {
        #region Field
        private readonly PinMapSheet _pinMapSheet;
        private readonly IoContiSheet _ioContiSheet;
        private readonly MultiTestSettingSheetsSingleton _testSettings = MultiTestSettingSheetsSingleton.Instance();
        private static readonly Regex _regex4 = new Regex(CommonConst.PinRegPattern, RegexOptions.Compiled);
        #endregion

        #region Constructor
        public GlobalSpecGenerator(PinMapSheet pinMapSheet, IoContiSheet ioContiSheet)
        {
            _pinMapSheet = pinMapSheet;
            _ioContiSheet = ioContiSheet;
        }
        #endregion

        #region Menber Function
        public List<GlobalSpec> GetPowerGlobalSpecs(PowerInfoSheet printSheet, MultiTestSettingSheetsSingleton allTestSettingData, IfoldPowerTableSheet ifoldPowerTable)
        {
            List<GlobalSpec> result = new List<GlobalSpec>();
            List<PowerInfoRow> powerInfos = printSheet.Rows;
            List<string> volTablePinList = allTestSettingData.PowerPinList;
            SplitDcvsDcviPowerPins(powerInfos);
            List<ProtocolAwarePin> nWirePin = NwireSingleton.Instance().SettingInfo.NwirePins;

            foreach (PowerInfoRow row in powerInfos)
            {
                string pinName = row.PinName;

                if (_regex4.IsMatch(pinName))
                {
                    pinName = _regex4.Match(pinName).Groups[CommonConst.OutFirst].ToString();
                }
                bool isNwirePin = NwireSingleton.Instance().SettingInfo.NwirePins.Exists(x => x.OutClk.Equals(pinName, StringComparison.OrdinalIgnoreCase));
                GlobalSpec powersequenceGlb;

                string ifoldValues = GetIfold(row, ifoldPowerTable);
                string powerPinIfoldCheckingComment = !volTablePinList.Contains(pinName) ? $"Missing pin in voltage table: {pinName}" : "";
                List<GlobalSpec> ifoldGlb = GetIfoldGlobalSpecs(pinName, ifoldValues, powerPinIfoldCheckingComment);
                if (powerPinIfoldCheckingComment != "" && !nWirePin.Any(x => pinName.Contains(x.OutClk)))
                {
                    ErrorReportManager.AddError(BasicErrorType.E_MissingPin_17, NeededSheets.PowerInfo, row.RowNum, 0, $"Missing pin in voltage table: {pinName}", new string[] { pinName }, new ErrorInfo() { Comments = new List<string>() { pinName } });
                }
                if (!isNwirePin)
                {
                    if (!row.IsIoPowerPin || !row.IsIoPowerDownPin)
                    {
                        if (!row.IsIoPowerPin)
                        {
                            AddPowerSequenceGlobalSpecs(result, pinName, row.PowerSequence, "PowerSequence");
                        }

                        if (!row.IsIoPowerDownPin)
                        {
                            AddPowerSequenceGlobalSpecs(result, pinName, row.PowerDownSequence, "PowerDownSequence");
                        }
                    }
                    else
                    {
                        powersequenceGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "PowerSequence", GlobalSpecSuffix), FormulaPrefix + GetPowerSeqValue(row.PowerSequence));
                        result.Add(powersequenceGlb);
                        GlobalSpec powerDownSequenceGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "PowerDownSequence", GlobalSpecSuffix), FormulaPrefix + GetPowerSeqValue(row.PowerDownSequence));
                        result.Add(powerDownSequenceGlb);
                    }
                }

                result.AddRange(ifoldGlb);

            }
            return result;
        }

        private void SplitDcvsDcviPowerPins(List<PowerInfoRow> powerInfos)
        {
            for (int i = 0; i < powerInfos.Count; i++)
            {
                string pinName = powerInfos[i].PinName;
                List<string> allPinType = PinService.GetPinTypeFromChannelMaps(pinName, _testSettings.PowerPinList, TestProgram.IgxlWorkBk.ChannelMapSheets);

                if (allPinType.Count > 1)
                {
                    IEnumerable<string> dcvspins = allPinType.Where(p => Regex.IsMatch(p, "DCVS", RegexOptions.IgnoreCase));
                    IEnumerable<string> dcvipins = allPinType.Where(p => Regex.IsMatch(p, "DCVI", RegexOptions.IgnoreCase));
                    if (dcvspins.Any() && dcvipins.Any())
                    {
                        if (allPinType.Count >= 2)
                        {
                            PowerInfoRow newDcvsPin = powerInfos[i].Copy();
                            newDcvsPin.PinName = pinName + "_DCVS";
                            powerInfos.Insert(i, newDcvsPin);
                            PowerInfoRow newDcviPin = powerInfos[i].Copy();
                            newDcviPin.PinName = pinName + "_DCVI";
                            powerInfos.Insert(i, newDcviPin);

                            powerInfos.RemoveAt(i + 2);
                        }
                    }
                }
            }
        }

        private List<GlobalSpec> GetIfoldGlobalSpecs(string pinName, string ifoldValues, string powerPinIfoldCheckingComment)
        {
            var ifoldGlb = new List<GlobalSpec>();
            foreach (string ifoldValue in ifoldValues.Split(';'))
            {
                string category = ifoldValue.Split(' ', '-', ':')[1].ToUpper();
                string keyword;
                switch (category)
                {
                    case "EVS":
                        keyword = SpecFormat.IfoldEvs;
                        break;
                    case "CONTI":
                        keyword = SpecFormat.IfoldConti;
                        break;
                    case "HIP":
                        keyword = SpecFormat.IfoldHip;
                        break;
                    default:
                        keyword = SpecFormat.Ifold;
                        break;
                }
                string finalPinName = JoinPinComponent(PinNamePrefix, pinName, keyword, GlobalSpecSuffix);
                GlobalSpec jobIfold = GetIFoldByJob(finalPinName, ifoldValue, powerPinIfoldCheckingComment);

                if (jobIfold != null)
                {
                    ifoldGlb.Add(jobIfold);
                }
            }
            return ifoldGlb;
        }

        private void AddPowerSequenceGlobalSpecs(List<GlobalSpec> result, string pinName, string sequenceValue, string sequenceName)
        {
            foreach (string item in sequenceValue.Split(';').ToList())
            {
                if (string.IsNullOrEmpty(item))
                {
                    continue;
                }

                string level = Regex.Match(item, @"t(?<num>\d+):(?<level>\w+)*", RegexOptions.IgnoreCase).Groups["level"].Value;
                string seq = Regex.Match(item, @"t(?<num>\d+):\s*", RegexOptions.IgnoreCase).Groups["num"].Value;
                GlobalSpec powersequenceGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, sequenceName, GlobalSpecSuffix, level), FormulaPrefix + GetPowerSeqValue(seq));
                result.Add(powersequenceGlb);
            }
        }

        private List<GlobalSpec> GetCommonIoPinsSpecs(IoInfoRow row)
        {
            string pinName = row.PinGrpName;  //Pins_0p6v                
            string valueNv = row.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(row.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : row.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "Ratio"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));  //Pins_0p6v_Vcl_Ratio
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", GlobalSpecSuffix), FormulaPrefix + (valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol));  //Pins_0p6v_Vcl_GLB
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "Ratio"), FormulaPrefix + "1.2");  //Pins_0p6v_Vch_Ratio
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", GlobalSpecSuffix), FormulaPrefix + (valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol));  //Pins_0p6v_Vch_GLB
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", GlobalSpecSuffix), FormulaPrefix + "0");  //Pins_0p6v_Ioh_GLB
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", GlobalSpecSuffix), FormulaPrefix + "0");  //Pins_0p6v_Iol_GLB
            GlobalSpec ratioVilGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", GlobalSpecSuffix), FormulaPrefix + row.VilRatio);  //Pins_0p6v_Ratio_Vil_GLB
            GlobalSpec vilGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", GlobalSpecSuffix), FormulaPrefix + (valueNv.Equals("0") ? "0" : "_" + pinGlb.Symbol + "*_" + ratioVilGlb.Symbol));  //Pins_0p6v_Vil_GLB
            GlobalSpec glbPlusCp1 = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix, "Plus"), FormulaPrefix + row.HvRatio);  //Pins_0p6v_GLB_Plus
            GlobalSpec glbMinusCp1 = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix, "Minus"), FormulaPrefix + row.LvRatio);  //Pins_0p6v_GLB_Minus
            GlobalSpec ratioVihGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", GlobalSpecSuffix), FormulaPrefix + row.VihRatio);  //Pins_0p6v_Ratio_Vih_GLB
            GlobalSpec vihGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", GlobalSpecSuffix), FormulaPrefix + (valueNv.Equals("0") ? "0" : "_" + pinGlb.Symbol + "*_" + ratioVihGlb.Symbol));  //Pins_0p6v_Vih_GLB
            GlobalSpec ratioVolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", GlobalSpecSuffix), FormulaPrefix + row.VolRatio);  //Pins_0p6v_Ratio_Vol_GLB
            GlobalSpec volGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", GlobalSpecSuffix), FormulaPrefix + (valueNv.Equals("0") ? "0" : "_" + pinGlb.Symbol + "*_" + ratioVolGlb.Symbol));  //Pins_0p6v_Vol_GLB
            GlobalSpec ratioVohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", GlobalSpecSuffix), FormulaPrefix + row.VohRatio);  //Pins_0p6v_Ratio_Voh_GLB
            GlobalSpec vohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", GlobalSpecSuffix), FormulaPrefix + (valueNv.Equals("0") ? "0" : "_" + pinGlb.Symbol + "*_" + ratioVohGlb.Symbol));  //Pins_0p6v_Voh_GLB
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!row.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(row.Vt, out _) ? row.Vt : "(_" + volGlb.Symbol + "+" + vohGlb.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", GlobalSpecSuffix), FormulaPrefix + row.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }

            GlobalSpec vtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);  //Pins_0p6v_Vt_GLB
            List<GlobalSpec> result = new List<GlobalSpec>
            {
                pinGlb,
                vclRatioGlb,
                vclGlb,
                vchRatioGlb,
                vchGlb,
                iohGlb,
                iolGlb,
                ratioVilGlb,
                vilGlb,
                glbPlusCp1,
                glbMinusCp1,
                ratioVihGlb,
                vihGlb,
                ratioVolGlb,
                volGlb,
                ratioVohGlb,
                vohGlb
            };
            if (row.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtGlb);
            return result;
        }

        private List<GlobalSpec> GetHardIpIoPins(IoInfoRow row, List<string> commonIoPins)
        {
            string pinName = row.PinGrpName;  //Pins_0p6v
            string valueNv = row.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(row.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : row.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec ratioVilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", "H", GlobalSpecSuffix), FormulaPrefix + row.VilRatio);
            GlobalSpec vilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVilH.Symbol);
            GlobalSpec glbPlusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "H", GlobalSpecSuffix, "Plus"), FormulaPrefix + row.HvRatio);
            GlobalSpec glbMinusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "H", GlobalSpecSuffix, "Minus"), FormulaPrefix + row.LvRatio);
            GlobalSpec ratioVihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", "H", GlobalSpecSuffix), FormulaPrefix + row.VihRatio);
            GlobalSpec vihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVihH.Symbol);
            GlobalSpec ratioVolH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", "H", GlobalSpecSuffix), FormulaPrefix + row.VolRatio);
            GlobalSpec volH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVolH.Symbol);
            GlobalSpec ratioVohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", "H", GlobalSpecSuffix), FormulaPrefix + row.VohRatio);
            GlobalSpec vohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVohH.Symbol);
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", "H", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", "H", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vcl", "H"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "H", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vcl) ? valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vcl));
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vch", "H"), FormulaPrefix + "1.2");
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "H", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vch) ? valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vch));
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!row.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(row.Vt, out double _) ? row.Vt : "(_" + volH.Symbol + "+" + vohH.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", "H", GlobalSpecSuffix), FormulaPrefix + row.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }
            GlobalSpec vtH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", "H", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);

            List<GlobalSpec> result = new List<GlobalSpec>();
            if (!commonIoPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(pinGlb);
            }
            result.Add(ratioVilH);
            result.Add(vilH);
            result.Add(glbPlusCp1H);
            result.Add(glbMinusCp1H);
            result.Add(ratioVihH);
            result.Add(vihH);
            result.Add(ratioVolH);
            result.Add(volH);
            result.Add(ratioVohH);
            result.Add(vohH);
            result.Add(iohGlb);
            result.Add(iolGlb);
            result.Add(vclRatioGlb);
            result.Add(vclGlb);
            result.Add(vchRatioGlb);
            result.Add(vchGlb);
            if (row.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtH);
            return result;
        }

        private List<GlobalSpec> GetScanIoPins(IoInfoRow row, List<string> commonIoPins)
        {
            string pinName = row.PinGrpName;  //Pins_0p6v
            string valueNv = row.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(row.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : row.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec ratioVilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", "Scan", GlobalSpecSuffix), FormulaPrefix + row.VilRatio);
            GlobalSpec vilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVilH.Symbol);
            GlobalSpec glbPlusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Scan", GlobalSpecSuffix, "Plus"), FormulaPrefix + row.HvRatio);
            GlobalSpec glbMinusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Scan", GlobalSpecSuffix, "Minus"), FormulaPrefix + row.LvRatio);
            GlobalSpec ratioVihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", "Scan", GlobalSpecSuffix), FormulaPrefix + row.VihRatio);
            GlobalSpec vihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVihH.Symbol);
            GlobalSpec ratioVolH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", "Scan", GlobalSpecSuffix), FormulaPrefix + row.VolRatio);
            GlobalSpec volH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVolH.Symbol);
            GlobalSpec ratioVohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", "Scan", GlobalSpecSuffix), FormulaPrefix + row.VohRatio);
            GlobalSpec vohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVohH.Symbol);
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", "Scan", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", "Scan", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vcl", "Scan"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "Scan", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vcl) ? valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vcl));
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vch", "Scan"), FormulaPrefix + "1.2");
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "Scan", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vch) ? valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vch));
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!row.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(row.Vt, out double _) ? row.Vt : "(_" + volH.Symbol + "+" + vohH.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", "Scan", GlobalSpecSuffix), FormulaPrefix + row.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }
            GlobalSpec vtH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", "Scan", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);

            List<GlobalSpec> result = new List<GlobalSpec>();
            if (!commonIoPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(pinGlb);
            }
            result.Add(ratioVilH);
            result.Add(vilH);
            result.Add(glbPlusCp1H);
            result.Add(glbMinusCp1H);
            result.Add(ratioVihH);
            result.Add(vihH);
            result.Add(ratioVolH);
            result.Add(volH);
            result.Add(ratioVohH);
            result.Add(vohH);
            result.Add(iohGlb);
            result.Add(iolGlb);
            result.Add(vclRatioGlb);
            result.Add(vclGlb);
            result.Add(vchRatioGlb);
            result.Add(vchGlb);
            if (row.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtH);
            return result;
        }

        private List<GlobalSpec> GetCountiIoPins(IoInfoRow row, List<string> commonIoPins)
        {
            string pinName = row.PinGrpName;  //Pins_0p6v
            string valueNv = row.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(row.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : row.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec ratioVilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VilRatio);
            GlobalSpec vilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVilH.Symbol);
            GlobalSpec glbPlusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Conti", GlobalSpecSuffix, "Plus"), FormulaPrefix + row.HvRatio);
            GlobalSpec glbMinusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Conti", GlobalSpecSuffix, "Minus"), FormulaPrefix + row.LvRatio);
            GlobalSpec ratioVihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VihRatio);
            GlobalSpec vihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVihH.Symbol);
            GlobalSpec ratioVolH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VolRatio);
            GlobalSpec volH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVolH.Symbol);
            GlobalSpec ratioVohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VohRatio);
            GlobalSpec vohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVohH.Symbol);
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", "Conti", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", "Conti", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vcl", "Conti"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "Conti", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vcl) ? valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vcl));
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vch", "Conti"), FormulaPrefix + "1.2");
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "Conti", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vch) ? valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vch));
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!row.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(row.Vt, out double _) ? row.Vt : "(_" + volH.Symbol + "+" + vohH.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }
            GlobalSpec vtH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", "Conti", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);

            List<GlobalSpec> result = new List<GlobalSpec>();
            if (!commonIoPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(pinGlb);
            }
            result.Add(ratioVilH);
            result.Add(vilH);
            result.Add(glbPlusCp1H);
            result.Add(glbMinusCp1H);
            result.Add(ratioVihH);
            result.Add(vihH);
            result.Add(ratioVolH);
            result.Add(volH);
            result.Add(ratioVohH);
            result.Add(vohH);
            result.Add(iohGlb);
            result.Add(iolGlb);
            result.Add(vclRatioGlb);
            result.Add(vclGlb);
            result.Add(vchRatioGlb);
            result.Add(vchGlb);
            if (row.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtH);
            return result;
        }

        private List<GlobalSpec> GetConcurrentHardIpPins(IoInfoRow row, List<string> commonIoPins)
        {
            string pinName = row.PinGrpName;  //Pins_0p6v
            string valueNv = row.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(row.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : row.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec ratioVilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", "H", GlobalSpecSuffix), FormulaPrefix + row.VilRatio);
            GlobalSpec vilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVilH.Symbol);
            GlobalSpec glbPlusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "H", GlobalSpecSuffix, "Plus"), FormulaPrefix + row.HvRatio);
            GlobalSpec glbMinusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "H", GlobalSpecSuffix, "Minus"), FormulaPrefix + row.LvRatio);
            GlobalSpec ratioVihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", "H", GlobalSpecSuffix), FormulaPrefix + row.VihRatio);
            GlobalSpec vihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVihH.Symbol);
            GlobalSpec ratioVolH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", "H", GlobalSpecSuffix), FormulaPrefix + row.VolRatio);
            GlobalSpec volH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVolH.Symbol);
            GlobalSpec ratioVohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", "H", GlobalSpecSuffix), FormulaPrefix + row.VohRatio);
            GlobalSpec vohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", "H", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVohH.Symbol);
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", "H", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", "H", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vcl", "H"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "H", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vcl) ? valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vcl));
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vch", "H"), FormulaPrefix + "1.2");
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "H", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vch) ? valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vch));
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!row.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(row.Vt, out double _) ? row.Vt : "(_" + volH.Symbol + "+" + vohH.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", "H", GlobalSpecSuffix), FormulaPrefix + row.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }
            GlobalSpec vtH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", "H", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);

            List<GlobalSpec> result = new List<GlobalSpec>();
            if (!commonIoPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(pinGlb);
            }
            result.Add(ratioVilH);
            result.Add(vilH);
            result.Add(glbPlusCp1H);
            result.Add(glbMinusCp1H);
            result.Add(ratioVihH);
            result.Add(vihH);
            result.Add(ratioVolH);
            result.Add(volH);
            result.Add(ratioVohH);
            result.Add(vohH);
            result.Add(iohGlb);
            result.Add(iolGlb);
            result.Add(vclRatioGlb);
            result.Add(vclGlb);
            result.Add(vchRatioGlb);
            result.Add(vchGlb);
            if (row.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtH);
            return result;
        }

        private List<GlobalSpec> GetConcurrentScanIoPins(IoInfoRow scanRow, List<string> commonIoPins)
        {
            string pinName = scanRow.PinGrpName;  //Pins_0p6v
            string valueNv = scanRow.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(scanRow.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : scanRow.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec ratioVilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", "Scan", GlobalSpecSuffix), FormulaPrefix + scanRow.VilRatio);
            GlobalSpec vilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVilH.Symbol);
            GlobalSpec glbPlusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Scan", GlobalSpecSuffix, "Plus"), FormulaPrefix + scanRow.HvRatio);
            GlobalSpec glbMinusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Scan", GlobalSpecSuffix, "Minus"), FormulaPrefix + scanRow.LvRatio);
            GlobalSpec ratioVihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", "Scan", GlobalSpecSuffix), FormulaPrefix + scanRow.VihRatio);
            GlobalSpec vihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVihH.Symbol);
            GlobalSpec ratioVolH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", "Scan", GlobalSpecSuffix), FormulaPrefix + scanRow.VolRatio);
            GlobalSpec volH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVolH.Symbol);
            GlobalSpec ratioVohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", "Scan", GlobalSpecSuffix), FormulaPrefix + scanRow.VohRatio);
            GlobalSpec vohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", "Scan", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVohH.Symbol);
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", "Scan", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", "Scan", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vcl", "Scan"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "Scan", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(scanRow.Vcl) ? valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol : scanRow.Vcl));
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vch", "Scan"), FormulaPrefix + "1.2");
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "Scan", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(scanRow.Vch) ? valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol : scanRow.Vch));
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!scanRow.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(scanRow.Vt, out double _) ? scanRow.Vt : "(_" + volH.Symbol + "+" + vohH.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", "Scan", GlobalSpecSuffix), FormulaPrefix + scanRow.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }
            GlobalSpec vtH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", "Scan", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);

            List<GlobalSpec> result = new List<GlobalSpec>();
            if (!commonIoPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(pinGlb);
            }
            result.Add(ratioVilH);
            result.Add(vilH);
            result.Add(glbPlusCp1H);
            result.Add(glbMinusCp1H);
            result.Add(ratioVihH);
            result.Add(vihH);
            result.Add(ratioVolH);
            result.Add(volH);
            result.Add(ratioVohH);
            result.Add(vohH);
            result.Add(iohGlb);
            result.Add(iolGlb);
            result.Add(vclRatioGlb);
            result.Add(vclGlb);
            result.Add(vchRatioGlb);
            result.Add(vchGlb);
            if (scanRow.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtH);
            return result;
        }

        private List<GlobalSpec> GetConcurrentCountiIoPins(IoInfoRow row, List<string> commonIoPins)
        {
            string pinName = row.PinGrpName;  //Pins_0p6v
            string valueNv = row.IsUsePowerPinValue ? IoInfoRow.GetIoPinGroupVoltage(row.PinGrpName.Trim(), _pinMapSheet, _ioContiSheet) : row.Nv;

            GlobalSpec pinGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, GlobalSpecSuffix), FormulaPrefix + valueNv);
            GlobalSpec ratioVilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vil", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VilRatio);
            GlobalSpec vilH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vil", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVilH.Symbol);
            GlobalSpec glbPlusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Conti", GlobalSpecSuffix, "Plus"), FormulaPrefix + row.HvRatio);
            GlobalSpec glbMinusCp1H = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Conti", GlobalSpecSuffix, "Minus"), FormulaPrefix + row.LvRatio);
            GlobalSpec ratioVihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vih", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VihRatio);
            GlobalSpec vihH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vih", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVihH.Symbol);
            GlobalSpec ratioVolH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vol", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VolRatio);
            GlobalSpec volH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vol", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVolH.Symbol);
            GlobalSpec ratioVohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Voh", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VohRatio);
            GlobalSpec vohH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Voh", "Conti", GlobalSpecSuffix), FormulaPrefix + "_" + pinGlb.Symbol + "*_" + ratioVohH.Symbol);
            GlobalSpec iohGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ioh", "Conti", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec iolGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Iol", "Conti", GlobalSpecSuffix), FormulaPrefix + "0");
            GlobalSpec vclRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vcl", "Conti"), FormulaPrefix + (valueNv.Equals("0") ? "0" : "-1/_" + pinGlb.Symbol));
            GlobalSpec vclGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vcl", "Conti", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vcl) ? valueNv.Equals("0") ? "0" : "_" + vclRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vcl));
            GlobalSpec vchRatioGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vch", "Conti"), FormulaPrefix + "1.2");
            GlobalSpec vchGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vch", "Conti", GlobalSpecSuffix), FormulaPrefix +
                                                (string.IsNullOrEmpty(row.Vch) ? valueNv.Equals("0") ? "0" : "_" + vchRatioGlb.Symbol + "*_" + pinGlb.Symbol : row.Vch));
            GlobalSpec ratioVtGlb = null;
            string vtGlbValue;
            if (!row.IsUsePowerPinValue)
            {
                vtGlbValue = double.TryParse(row.Vt, out double _) ? row.Vt : "(_" + volH.Symbol + "+" + vohH.Symbol + ")*0.5";
            }
            else
            {
                ratioVtGlb = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Ratio", "Vt", "Conti", GlobalSpecSuffix), FormulaPrefix + row.VtRatio);  //Pins_0p6v_Ratio_Vt_GLB
                vtGlbValue = "_" + pinGlb.Symbol + "*_" + ratioVtGlb.Symbol;
            }
            GlobalSpec vtH = new GlobalSpec(JoinPinComponent(PinNamePrefix, pinName, "Vt", "Conti", GlobalSpecSuffix), FormulaPrefix + vtGlbValue);

            List<GlobalSpec> result = new List<GlobalSpec>();
            if (!commonIoPins.Exists(x => x.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(pinGlb);
            }
            result.Add(ratioVilH);
            result.Add(vilH);
            result.Add(glbPlusCp1H);
            result.Add(glbMinusCp1H);
            result.Add(ratioVihH);
            result.Add(vihH);
            result.Add(ratioVolH);
            result.Add(volH);
            result.Add(ratioVohH);
            result.Add(vohH);
            result.Add(iohGlb);
            result.Add(iolGlb);
            result.Add(vclRatioGlb);
            result.Add(vclGlb);
            result.Add(vchRatioGlb);
            result.Add(vchGlb);
            if (row.IsUsePowerPinValue)
            {
                result.Add(ratioVtGlb);
            }

            result.Add(vtH);
            return result;
        }

        public List<GlobalSpec> GetIoGlobalSpecs(IoInfoSheet printData, IoInfoSheet printConcurrentData)
        {
            List<GlobalSpec> result = new List<GlobalSpec>();
            List<IoInfoRow> ioInfos = printData.GetBlockIoInfo("Common");
            List<IoInfoRow> hardipIoInfos = printData.GetBlockIoInfo("HardIP");
            List<IoInfoRow> scanIoInfos = printData.GetBlockIoInfo("Scan");
            List<IoInfoRow> contiIoInfos = printData.GetBlockIoInfo("DCTEST_Continuity");
            List<string> commonIoPins = printData.GetBlockPins("Common").ToList();

            //common io pins
            foreach (IoInfoRow row in ioInfos)
            {
                result.AddRange(GetCommonIoPinsSpecs(row));
            }

            //hardip io pins
            foreach (IoInfoRow hardipRow in hardipIoInfos)
            {
                result.AddRange(GetHardIpIoPins(hardipRow, commonIoPins));
            }

            //scan io pins
            foreach (IoInfoRow scanRow in scanIoInfos)
            {
                result.AddRange(GetScanIoPins(scanRow, commonIoPins));
            }

            //conti io pins
            foreach (IoInfoRow contiRow in contiIoInfos)
            {
                result.AddRange(GetCountiIoPins(contiRow, commonIoPins));
            }

            // Start handling concurrent
            List<string> commonIoPinsConcurrent =
                printConcurrentData != null && printConcurrentData.GetBlockIoInfo("Common").Any()
                    ? printConcurrentData.GetBlockPins("Common").ToList()
                    : new List<string>();

            List<IoInfoRow> hardipIoInfosConcurrent =
                printConcurrentData != null && printConcurrentData.GetBlockIoInfo("HardIP").Any()
                    ? printConcurrentData.GetBlockIoInfo("HardIP")
                    : new List<IoInfoRow>();
            //hardip io pins
            foreach (IoInfoRow hardipRow in hardipIoInfosConcurrent)
            {
                result.AddRange(GetConcurrentHardIpPins(hardipRow, commonIoPinsConcurrent));
            }

            List<IoInfoRow> scanIoInfosConcurrent =
                printConcurrentData != null && printConcurrentData.GetBlockIoInfo("Scan").Any()
                    ? printConcurrentData.GetBlockIoInfo("Scan")
                    : new List<IoInfoRow>();
            //scan io pins
            foreach (IoInfoRow scanRow in scanIoInfosConcurrent)
            {
                result.AddRange(GetConcurrentScanIoPins(scanRow, commonIoPinsConcurrent));
            }

            List<IoInfoRow> contiIoInfosConcurrent =
                printConcurrentData != null
                && printConcurrentData.GetBlockIoInfo("DCTEST_Continuity").Any()
                    ? printConcurrentData.GetBlockIoInfo("DCTEST_Continuity")
                    : new List<IoInfoRow>();
            //conti io pins
            foreach (IoInfoRow contiRow in contiIoInfosConcurrent)
            {
                result.AddRange(GetConcurrentCountiIoPins(contiRow, commonIoPinsConcurrent));
            }

            return result;
        }

        internal string GetIfold(PowerInfoRow powerInfoRow, IfoldPowerTableSheet ifoldPowerTableSheet)
        {
            IfoldPowerTableRow targetIfold = ifoldPowerTableSheet?.Rows.Find(x => x.PinName.Equals(powerInfoRow.PinName, StringComparison.OrdinalIgnoreCase));
            if (targetIfold != null)
            {
                return targetIfold.Current;
            }
            var ifolds = new List<string>();
            foreach (string ifoldJob in powerInfoRow.Ifold.Split(';'))
            {
                if (ifoldJob.Contains(":"))
                {
                    if (string.IsNullOrEmpty(ifoldJob.Split(':')[0]))
                    {
                        ifolds.Add($"0:{ifoldJob.Split(':')[1]}");
                    }
                    else
                    {
                        ifolds.Add(ifoldJob);
                    }
                }
                else if (!string.IsNullOrEmpty(ifoldJob))
                {
                    ifolds.Add(ifoldJob);
                }
                else if (double.TryParse(ifoldJob, out double _))
                {
                    ifolds.Add(ifoldJob);
                }
                else
                {
                    ifolds.Add("0");
                }
            }
            powerInfoRow.Ifold = string.Join(";", ifolds);
            return powerInfoRow.Ifold;
        }

        internal string GetPowerSeqValue(string strSequenceValue)
        {
            string lStrResult = "";
            string lStrMatchPattern = @"^[a-zA-Z]*(?<str>\d+)";
            if (Regex.IsMatch(strSequenceValue, lStrMatchPattern))
            {
                lStrResult = Regex.Match(strSequenceValue, lStrMatchPattern).Groups["str"].ToString();
                if (int.TryParse(lStrResult, out int value))
                {
                    return value.ToString();
                }
            }
            else if (string.IsNullOrEmpty(strSequenceValue))
            {

            }

            else
            {
                lStrResult = "99";
            }
            return lStrResult;
        }

        private GlobalSpec GetIFoldByJob(string name, string value, string glbComment)
        {
            string[] glbSgmts = value.Split(':');
            var spec = new GlobalSpec(name, FormulaPrefix + glbSgmts[0])
            {
                Comment = glbComment
            };
            string job = "";
            List<string> jobList = LocalSpecs.AllJobs;
            if (glbSgmts.Length > 1)
            {
                if (glbSgmts[1].ContainsIgnoreCase("EVS") || glbSgmts[1].ContainsIgnoreCase("CONTI") || glbSgmts[1].ContainsIgnoreCase("HIP"))
                {
                    string[] jobs = glbSgmts[1].Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
                    if (jobs.Any() && jobs.Length > 1)
                    {
                        job = jobs.Last();
                    }
                }
                else
                {
                    job = glbSgmts[1].Trim();
                }
            }
            if (!string.IsNullOrEmpty(job))
            {
                if (jobList.Any(x => x.Equals(job, StringComparison.OrdinalIgnoreCase)))
                {
                    spec.Job = job;
                    return spec;
                }
            }
            return null;
        }
        #endregion
    }
}
