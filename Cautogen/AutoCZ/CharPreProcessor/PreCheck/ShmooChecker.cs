using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.ReportManager;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using CommonLib.ErrorReport;
using CommonLib.ErrorReport.Base;
using CommonLib.ErrorReport.ErrorCodes;

namespace Cautogen.AutoCZ.CharPreProcessor.PreCheck
{
    public class ShmooChecker : PreCheckBase
    {
        public override void Check(List<Characterization> charList, string sheetName)
        {
            foreach (Characterization charRow in charList)
            {
                var shmooList = new List<ShmooSpec>();
                shmooList.AddRange(charRow.PowerSupplyX);
                shmooList.AddRange(charRow.PowerSupplyY);
                shmooList.AddRange(charRow.PowerSupplyZ);

                // shmoo points should be integer and step >= 6
                _CheckShmooPoints(shmooList, charRow);

                _CheckUserdef3MatchesShmooSetting(charRow, shmooList);

                _CheckMissingShmooCondition(charRow, shmooList);

                _CheckNv(shmooList, charRow);

                _CheckMissingPinName(shmooList, charRow);

                _CheckTrackingPinStepMismatchPrimary(charRow, shmooList);

                _CheckTrackingValtConsist(charRow, shmooList);

                _CheckShmooSramPins(shmooList, charRow);

                _CheckVoltageMoreThan1p3(shmooList, charRow);

                _CheckShmooLowToHigh(shmooList, charRow);
            }
        }

        private static void _CheckShmooLowToHigh(IEnumerable<ShmooSpec> shmooList, Characterization charRow)
        {
            foreach (ShmooSpec shmoo in shmooList)
            {
                if (Regex.IsMatch(shmoo.Name, "^vdd", RegexOptions.IgnoreCase))
                {
                    _ = new List<int>();
                    if (decimal.TryParse(shmoo.Start, out decimal shmooStartValue) && decimal.TryParse(shmoo.Stop, out decimal shmooStopValue))
                    {
                        if (shmooStartValue < shmooStopValue)
                        {
                            string outString = "Shmoo range is from low to high";
                            ErrorManager.AddWarning(ErrorType.VddShmooLowToHigh, charRow, new List<int> { shmoo.ColIndex }, outString);
                            ErrorReportManager.AddError(CharErrorType.E_VddShmooLowToHigh_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, []);
                        }
                    }
                }
            }
        }

        private static void _CheckVoltageMoreThan1p3(IEnumerable<ShmooSpec> shmooList, Characterization charRow)
        {
            foreach (ShmooSpec shmoo in shmooList)
            {
                if (Regex.IsMatch(shmoo.Name, "^vdd", RegexOptions.IgnoreCase))
                {
                    _ = new List<int>();
                    if (decimal.TryParse(shmoo.Start, out decimal shmooValue))
                    {
                        if (Math.Abs(shmooValue) >= (decimal)1.3)
                        {
                            string outString = "Shmoo start value >= 1.3";
                            ErrorManager.AddWarning(ErrorType.VoltageHigherThan1P3, charRow, new List<int> { shmoo.ColIndex }, outString);
                            ErrorReportManager.AddError(CharErrorType.W_VoltageHigherThan1P3_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, ["start"]);
                        }
                    }
                    if (decimal.TryParse(shmoo.Stop, out shmooValue))
                    {
                        if (Math.Abs(shmooValue) >= (decimal)1.3)
                        {
                            string outString = "Shmoo stop value >= 1.3";
                            ErrorManager.AddWarning(ErrorType.VoltageHigherThan1P3, charRow, new List<int> { shmoo.ColIndex }, outString);
                            ErrorReportManager.AddError(CharErrorType.W_VoltageHigherThan1P3_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, ["stop"]);
                        }
                    }
                }
            }
        }

        private static void _CheckShmooSramPins(IList<ShmooSpec> shmooList, Characterization item)
        {
            Dictionary<string, List<string>> selsramDic = GenselsramDic(SelsramTableReader.SelSramRows);
            var charstep = shmooList.Where(x => !string.IsNullOrEmpty(x.Step) && x.Step != "0").Select(x => x.Name).ToList();

            foreach (string charpin in charstep)
            {
                if (selsramDic.ContainsKey(charpin))
                {
                    if (!selsramDic[charpin].All(charstep.Contains))
                    {
                        ErrorManager.Add(new ErrorMessage
                        {
                            ErrorLevel = ErrorLevel.Warning,
                            ErrorType = ErrorType.SramPinNotTieLogic,
                            SheetName = item.SheetName,
                            RowNum = item.RowNum,
                            Message = string.Format("Shmoo pin only define: " + charpin + "; suggest to tie:" + charpin + "," + string.Join(",", selsramDic[charpin])),
                            CommentsList = new List<string> { },
                        });
                        ErrorReportManager.AddError(CharErrorType.W_SramPinNotTieLogic_01, item.SheetName, item.RowNum, 0,
                            [charpin, charpin, string.Join(",", selsramDic[charpin])],
                            new ErrorInfo() { Comments = new List<string> { } });
                    }
                }
            }
        }

        private static Dictionary<string, List<string>> GenselsramDic(List<SelSrmRow> selsrmMappingTalbeRows)
        {
            var selsramDic = new Dictionary<string, List<string>>();
            foreach (SelSrmRow selsrmRow in selsrmMappingTalbeRows)
            {
                var sramLogicList = new List<string>();
                if (!selsramDic.ContainsKey(selsrmRow.SramPins.Replace("_", "")))
                {
                    sramLogicList.Add(selsrmRow.LogicPins.Replace("_", ""));
                    selsramDic.Add(selsrmRow.SramPins.Replace("_", ""), sramLogicList);
                }
                else
                {
                    sramLogicList = selsramDic[selsrmRow.SramPins.Replace("_", "")];
                    sramLogicList.Add(selsrmRow.LogicPins.Replace("_", ""));
                    selsramDic[selsrmRow.SramPins.Replace("_", "")] = sramLogicList.Distinct().ToList();
                }

            }
            return selsramDic;
        }

        private static void _CheckMissingPinName(IEnumerable<ShmooSpec> shmooList, Characterization charRow)
        {
            // bypass check if user does not input pin map;
            if (UtilityMain.UtilityData.InputParam.PinMapFile == "")
            {
                if (!UtilityMain.UtilityData.InputParam.CharPreCheckWithoutTp)
                {
                    var missingPinMap = new ErrorMessage
                    {
                        Message = "Missing PinMap.txt Please Reload PinMap or Test Program!!!",
                        ErrorType = ErrorType.MissingPinMap,
                        ErrorLevel = ErrorLevel.Error
                    };
                    ErrorManager.Add(missingPinMap);
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinMap_01, "Error", 0, 0, []);
                }
                return;
            }
            foreach (ShmooSpec shmoo in shmooList)
            {
                if (Regex.IsMatch(shmoo.Name, "^vdd", RegexOptions.IgnoreCase) && !UtilityMain.UtilityFunction.PinExistInPinMap(shmoo.Name))
                {
                    const string outString = "Missing VDD name in PinMap.txt";
                    ErrorManager.AddError(ErrorType.MissingPinName, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, "Use", outString, shmoo.Name);
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinName_03, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { shmoo.Name } });
                }
            }
        }

        private static void _CheckNv(List<ShmooSpec> shmooList, Characterization charRow)
        {
            if (!UtilityMain.UtilityData.InputParam.NvIsChecked)
            {
                return;
            }

            // Check shmoo range with Global Specs
            foreach (ShmooSpec shmoo in shmooList.Where(shmoo => shmoo.Start != "" && shmoo.Stop != "" && shmoo.Start != shmoo.Stop))
            {
                string key;
                if (charRow.Group == "P0" || !_CheckMode(charRow, shmoo.Name))
                {
                    key = shmoo.Name.ToUpper();
                }
                else
                {
                    key = (shmoo.Name + charRow.Group).ToUpper();
                }
                //string key = shmoo.Name.ToUpper();
                if (GlobalSpecsReader.GlobalSpecs.TryGetValue(key, out double nv))
                {
                    double start = Convert.ToDouble(shmoo.Start);
                    double stop = Convert.ToDouble(shmoo.Stop);
                    if (!((start - nv) * (stop - nv) >= 0))
                    {
                        continue;
                    }

                    const string outString = "NV GlobalSpecs not within Start/Stop";
                    ErrorManager.AddError(ErrorType.WrongShmooNv, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, charRow.Use,
                        outString, shmoo.Name);
                    ErrorReportManager.AddError(CharErrorType.E_WrongShmooNv_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { shmoo.Name } });
                }
                else
                {
                    const string outString = "Shmoo pin does not exist in GlobalSpecs";
                    ErrorManager.AddError(ErrorType.MissingPinInGlobalSpecs, charRow.SheetName, charRow.RowNum, shmoo.ColIndex,
                        charRow.Use, outString, shmoo.Name);
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinInGlobalSpecs_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { shmoo.Name } });
                }
            }

            // Check power condition with rule equation
            if (!CharPlan.HardIpSheets.Contains(charRow.SheetName.ToUpper()))
            {
                var shmooStart = new Hashtable();
                var shmooStop = new Hashtable();
                foreach (ShmooSpec shmoo in shmooList)
                {
                    shmooStart.Add(shmoo.Name, shmoo.Start);
                    shmooStop.Add(shmoo.Name, shmoo.Stop);
                }
            }

            // Check shmoo range within NV * 1.7
            foreach (ShmooSpec shmoo in shmooList.Where(shmoo => shmoo.Start != ""))
            {
                string key;
                if (charRow.Group == "P0" || !_CheckMode(charRow, shmoo.Name))
                {
                    key = shmoo.Name.ToUpper();
                }
                else
                {
                    key = (shmoo.Name + charRow.Group).ToUpper();
                }

                if (GlobalSpecsReader.GlobalSpecs.TryGetValue(key, out double nv))
                {
                    double start = Convert.ToDouble(shmoo.Start);
                    double stop = Convert.ToDouble(shmoo.Stop);
                    if (!(Math.Abs(start) > Math.Abs(nv) * 1.7) && !(Math.Abs(stop) > Math.Abs(nv) * 1.7))
                    {
                        continue;
                    }

                    const string outString = "VDD power setting over than 1.7 of  nominal voltage ";
                    ErrorManager.AddError(ErrorType.WrongShmooNv, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, charRow.Use, outString,
                        shmoo.Name);
                    ErrorReportManager.AddError(CharErrorType.E_WrongShmooNv_02, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { shmoo.Name } });
                }
                else
                {
                    const string outString = "Shmoo pin does not exist in GlobalSpecs";
                    ErrorManager.AddError(ErrorType.MissingPinInGlobalSpecs, charRow.SheetName, charRow.RowNum,
                        shmoo.ColIndex, charRow.Use, outString, shmoo.Name);
                    ErrorReportManager.AddError(CharErrorType.E_MissingPinInGlobalSpecs_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, [],
                        new ErrorInfo() { Comments = new List<string>() { shmoo.Name } });
                }
            }
        }

        private static void _CheckMissingShmooCondition(Characterization charRow, IEnumerable<ShmooSpec> shmooList)
        {
            // reset previous condition
            bool hasStartStopInPower = AnyPowerSupplyHasShmooRange(charRow);

            bool hasStartStopVal = _CheckStartStopHasValue(shmooList, charRow) || AnyPinHasShmooRange(charRow.Pins);

            // should have start stop value for user def1 is functional header
            if (!hasStartStopVal && charRow.IsFuncRow())
            {
                string outString = "Missing shmoo condition in " + charRow.SheetName + "  Row: " + charRow.RowNum;
                ErrorManager.AddError(ErrorType.MissingShmooCondition, charRow.SheetName, charRow.RowNum, shmooList.Select(p => p.ColIndex).ToList(), charRow.Use, outString);
                foreach (int col in shmooList.Select(p => p.ColIndex).ToList())
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingShmooCondition_01, charRow.SheetName, charRow.RowNum, col, [charRow.SheetName, $"{charRow.RowNum}"]);
                }
            }

            // for HIO, should have some pin sweep setting
            else if (charRow.UserDef1.ToUpper() == "HIO" && charRow.Pins.Count == 0)
            {
                string outString = "Missing VIH/VIL shmoo for HIO in " + charRow.SheetName + "  Row: " + charRow.RowNum;
                ErrorManager.AddError(ErrorType.MissingShmooCondition, charRow.SheetName, charRow.RowNum,
                    charRow.ColNum("userdef1"), charRow.Use, outString);
                foreach (int col in charRow.ColNum("userdef1"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_MissingShmooCondition_02, charRow.SheetName, charRow.RowNum, col, [charRow.SheetName, $"{charRow.RowNum}"]);
                }
            }

            // for HAC, should not have start != stop in power/io pin in normal case
            else if (charRow.UserDef1.ToUpper() == "HAC" && (hasStartStopVal || hasStartStopInPower))
            {
                if (IsSweepCodeOrVoltage(charRow))
                {
                    return;
                }

                string outString = "Dummy shmoo condition for HAC in " + charRow.SheetName + "  Row: " + charRow.RowNum;
                ErrorManager.AddWarning(ErrorType.DummyShmooCondition, charRow.SheetName, charRow.RowNum, charRow.Pins.Select(p => p.PlanIndex).ToList(), charRow.Use, outString);
                foreach (int col in charRow.Pins.Select(p => p.PlanIndex).ToList())
                {
                    ErrorReportManager.AddError(CharErrorType.W_DummyShmooCondition_01, charRow.SheetName, charRow.RowNum, col, [charRow.SheetName, $"{charRow.RowNum}"]);
                }
            }
        }

        private static bool AnyPowerSupplyHasShmooRange(Characterization charRow)
        {
            return AnyShmooSpecHasShmooRange(charRow.PowerSupplyX) ||
                   AnyShmooSpecHasShmooRange(charRow.PowerSupplyY) ||
                   AnyShmooSpecHasShmooRange(charRow.PowerSupplyZ);
        }

        private static bool AnyShmooSpecHasShmooRange(IEnumerable<ShmooSpec> specs)
        {
            return specs.Any(s => s.Start != "" && s.Stop != "" && s.Start != s.Stop);
        }

        private static bool AnyPinHasShmooRange(IEnumerable<Pin> pins)
        {
            return pins.Any(pin => pin.Start != "" && pin.Stop != "" && pin.Start != pin.Stop);
        }

        private static bool IsSweepCodeOrVoltage(Characterization charRow)
        {
            return charRow.UserDef8.Contains("SweepCode") || charRow.UserDef8.Contains("SweepVoltage");
        }

        private static void _CheckUserdef3MatchesShmooSetting(Characterization charRow, List<ShmooSpec> shmooList)
        {
            if (!charRow.IsFuncRow())
            {
                return;
            }

            var shmoopin = shmooList.Where(a => Regex.IsMatch(a.Name, charRow.UserDef3)).ToList();
            var colorRow = new List<int>();
            if (shmoopin.Any())
            {
                foreach (ShmooSpec item in shmoopin)
                {
                    for (int i = 0; i < 3; i++)
                    {
                        colorRow.Add(item.ColIndex + i);
                    }

                    if (item.IsSweep)
                    {
                        return;
                    }
                }
                colorRow.AddRange(charRow.ColNum("userdef3"));
                ErrorManager.AddError(ErrorType.Userdef3MismatchShmoo, charRow.SheetName, charRow.RowNum,
                    colorRow, charRow.Use,
                    "There is an issue with USERDEF3 related Vmain or Valt pin shmoo setting");
                foreach (int col in colorRow)
                {
                    ErrorReportManager.AddError(CharErrorType.E_Userdef3MismatchShmoo_01, charRow.SheetName, charRow.RowNum, col, []);
                }

            }
            else//(shmoo == null)
            {
                ShmooSpec shmoo =
                    shmooList.FirstOrDefault(a => a.Name.Equals("VDD" + Regex.Replace(charRow.UserDef3, "^VDD", "")));
                if (shmoo != null && shmoo.Start != "" && shmoo.Stop != "" && shmoo.Start != shmoo.Stop)
                {
                    return;
                }

                if (shmoo == null) //Added on 2016/11/30. Pin Sweep may be "XI0" and defined in Pins
                {
                    Pin pin = charRow.Pins.FirstOrDefault(a => a.PinName.Equals(charRow.UserDef3));
                    if (pin != null && pin.Start != "" && pin.Stop != "" && pin.Start != pin.Stop)
                    {
                        return;
                    }
                }
                ErrorManager.AddError(ErrorType.Userdef3MismatchShmoo, charRow.SheetName, charRow.RowNum,
                    charRow.ColNum("userdef3"), charRow.Use,
                    "USERDEF3 power does not match with shmoo setting");
                foreach (int col in charRow.ColNum("userdef3"))
                {
                    ErrorReportManager.AddError(CharErrorType.E_Userdef3MismatchShmoo_02, charRow.SheetName, charRow.RowNum, col, []);
                }
            }

        }

        private static bool _CheckShmooIsNumeric(ShmooSpec spec, Characterization item)
        {
            bool result = true;
            var nonNumCol = new List<int>();

            if (!decimal.TryParse(spec.Start, out _) && !string.IsNullOrEmpty(spec.Start))
            {
                result = false;
                nonNumCol.Add(spec.ColIndex);
            }

            if (!decimal.TryParse(spec.Stop, out _) && !string.IsNullOrEmpty(spec.Stop))
            {
                result = false;
                nonNumCol.Add(spec.ColIndex + 1);
            }

            if (!decimal.TryParse(spec.Step, out _) && !string.IsNullOrEmpty(spec.Step))
            {
                result = false;
                nonNumCol.Add(spec.ColIndex + 2);
            }

            if (nonNumCol.Count > 0)
            {
                ErrorManager.AddError(ErrorType.IllegalShmooValue, item, nonNumCol, "Shmoo value is non-numeric string");
                foreach (int col in nonNumCol)
                {
                    ErrorReportManager.AddError(CharErrorType.E_IllegalShmooValue_01, item.SheetName, item.RowNum, col, []);
                }
            }

            return result;
        }

        private static void _CheckShmooPoints(IList<ShmooSpec> shmooList, Characterization item)
        {
            foreach (ShmooSpec spec in shmooList)
            {
                if (!_CheckShmooIsNumeric(spec, item))
                {
                    continue;
                }

                double stepVal;
                if (spec.Start == "" || spec.Stop == "" || spec.Start == spec.Stop)
                {
                    continue;
                }

                if (spec.Start != spec.Stop && spec.Step != "" && double.Parse(spec.Step) == 0)
                {
                    ErrorManager.AddError(ErrorType.WrongShmooSteps, item.SheetName, item.RowNum, spec.ColIndex, item.Use, "Start didn't equal Stop but step is 0");
                    ErrorReportManager.AddError(CharErrorType.E_WrongShmooSteps_01, item.SheetName, item.RowNum, spec.ColIndex, []);
                    continue;
                }

                if (spec.Step != "")
                {
                    double.TryParse(spec.Step, out stepVal);
                }
                else
                {
                    stepVal = 0.005;
                }

                if (stepVal > 0.005)
                {
                    ErrorManager.AddWarning(ErrorType.WrongShmooSteps, item, new List<int> { spec.ColIndex + 2 }, "Shmoo step > 0.005");
                    foreach (int col in new List<int> { spec.ColIndex + 2 })
                    {
                        ErrorReportManager.AddError(CharErrorType.W_WrongShmooSteps_02, item.SheetName, item.RowNum, col, []);
                    }
                }

                decimal points = (Convert.ToDecimal(spec.Start) - Convert.ToDecimal(spec.Stop)) / Convert.ToDecimal(stepVal);

                if (Convert.ToDecimal(spec.Start) < Convert.ToDecimal(spec.Stop))
                {
                    ErrorManager.AddWarning(ErrorType.HeatAlarm, item, new List<int> { spec.ColIndex, spec.ColIndex + 1 },
                        "Stop voltage higher than start voltage. It might cause heat alarm, please check.");
                    foreach (int col in new List<int> { spec.ColIndex, spec.ColIndex + 1 })
                    {
                        ErrorReportManager.AddError(CharErrorType.W_HeatAlarm_01, item.SheetName, item.RowNum, col, []);
                    }
                }

                if (Convert.ToInt32(points) != points)
                {
                    ErrorManager.AddWarning(ErrorType.WrongShmooSteps, item, new List<int> { spec.ColIndex }, "Shmoo point is not integer");
                    ErrorReportManager.AddError(CharErrorType.W_WrongShmooSteps_03, item.SheetName, item.RowNum, spec.ColIndex, []);
                }

                if (Math.Abs(points) < 6)
                {
                    ErrorManager.AddError(ErrorType.LessThan6ShmooPoints, item.SheetName, item.RowNum, spec.ColIndex, item.Use, "Less than 6sShmoo points");
                    ErrorReportManager.AddError(CharErrorType.E_LessThan6ShmooPoints_01, item.SheetName, item.RowNum, spec.ColIndex, []);
                }
            }
        }

        private static bool _CheckStartStopHasValue(IEnumerable<ShmooSpec> shmooList, Characterization item)
        {
            bool hasStartStopVal = false;
            string missingShmooValue = "Missing start or stop of: ";
            var errIndX = new List<int>();
            foreach (ShmooSpec shmoo in shmooList)
            {
                if (!(shmoo.IsAllBlank && shmoo.Name != item.UserDef3) &&
                    (!double.TryParse(shmoo.Start, out double _) || !double.TryParse(shmoo.Stop, out double _)))
                {
                    missingShmooValue += shmoo.Name;
                    errIndX.Add(shmoo.ColIndex);
                }
                else if ((shmoo.Start != "" && shmoo.Stop == "") || (shmoo.Start == "" && shmoo.Stop != ""))
                {
                    missingShmooValue += shmoo.Name;
                    errIndX.Add(shmoo.ColIndex);

                }
                else if (shmoo.Start != "" && shmoo.Stop != "" && shmoo.Start != shmoo.Stop)
                {
                    hasStartStopVal = true;
                }
            }

            if (missingShmooValue == "Missing start or stop of: ")
            {
                return hasStartStopVal;
            }

            string outString = missingShmooValue;
            ErrorManager.AddError(ErrorType.MissingShmooValue, item.SheetName, item.RowNum, errIndX, item.Use, outString);
            foreach (int col in errIndX)
            {
                ErrorReportManager.AddError(CharErrorType.E_MissingShmooValue_01, item.SheetName, item.RowNum, col, [outString]);
            }

            return hasStartStopVal;
        }

        private static bool _CheckMode(Characterization item, string shmooName)
        {
            return (Regex.IsMatch(item.Group, "^MC.*") && shmooName.Contains("CPU")) ||
                   (Regex.IsMatch(item.Group, "^MG.*") && shmooName.Contains("GPU")) ||
                   (Regex.IsMatch(item.Group, "^MS.*") && shmooName.Contains("SOC"));
        }

        private static void _CheckTrackingPinStepMismatchPrimary(Characterization charRow, List<ShmooSpec> shmooList)
        {
            string mismatchTrackingStep = "Mismatch tracking pins: ";
            if (!charRow.IsFuncRow())
            {
                return;
            }

            ShmooSpec shmoo = shmooList.FirstOrDefault(a => a.Name.Equals(charRow.UserDef3));
            if (shmoo != null && shmoo.Start != "" && shmoo.Stop != "" && shmoo.Start != shmoo.Stop)
            {
                foreach (ShmooSpec sl in shmooList)
                {
                    if (shmoo.Name == sl.Name)
                    {
                        continue;
                    }

                    if (sl.Start != "" && sl.Stop != "" && sl.Start != sl.Stop && sl.Step != "" && shmoo.Step != "")
                    {
                        if (double.Parse(sl.Step) == 0)
                        {
                            continue;    //check in _CheckShmooPoints()
                        }

                        if (shmoo.Start != sl.Start || shmoo.Stop != sl.Stop)
                        {
                            float shmooStepSize = Math.Abs(float.Parse(shmoo.Start) - float.Parse(shmoo.Stop)) / float.Parse(shmoo.Step);
                            float slStepSize = Math.Abs(float.Parse(sl.Start) - float.Parse(sl.Stop)) / float.Parse(sl.Step);
                            if (shmooStepSize != slStepSize)
                            {
                                if (sl.Order == shmoo.Order && sl.Order != "" && shmoo.Order != "")
                                {
                                    mismatchTrackingStep += sl.Name + ",";
                                }
                            }
                        }
                    }
                }
            }

            if (mismatchTrackingStep != "Mismatch tracking pins: ")
            {
                string outString = mismatchTrackingStep.Substring(0, mismatchTrackingStep.Length - 1);
                ErrorManager.AddError(ErrorType.TrackingPinMismatchPrimaryStep, charRow.SheetName, charRow.RowNum, shmoo.ColIndex, charRow.Use, outString);
                ErrorReportManager.AddError(CharErrorType.E_TrackingPinMismatchPrimaryStep_01, charRow.SheetName, charRow.RowNum, shmoo.ColIndex,
                    [mismatchTrackingStep.Substring(mismatchTrackingStep.IndexOf(':') + 1)]);
            }
        }

        private static void _CheckTrackingValtConsist(Characterization charRow, IEnumerable<ShmooSpec> shmooList)
        {
            if (!charRow.IsFuncRow())
            {
                return;
            }

            IEnumerable<IGrouping<string, ShmooSpec>> shmooSetsByOrder =
                shmooList.Where(shmoo => shmoo.Start != "" && shmoo.Stop != "" && shmoo.Start != shmoo.Stop)
                .GroupBy(p => p.Order);

            var errorCol = new List<int>();
            foreach (IGrouping<string, ShmooSpec> shmooSets in shmooSetsByOrder)
            {

                if (shmooSets.Count() > 1 && shmooSets.ToList().Exists(shmoo => shmoo.IsValt))
                {
                    if (shmooSets.Count(p => p.IsValt) != shmooSets.Count())
                    {
                        errorCol.AddRange(shmooSets.Select(p => p.ColIndex));

                    }
                }
            }
            if (errorCol.Count > 0)
            {
                const string errMsg = "Shmoo Setting Contains Valt and Vmain";
                ErrorManager.AddError(ErrorType.PrimaryTrackingPinValtInconsist,
                    charRow.SheetName, charRow.RowNum, errorCol, charRow.Use, errMsg);
                foreach (int col in errorCol)
                {
                    ErrorReportManager.AddError(CharErrorType.E_PrimaryTrackingPinValtInconsist_01,
                        charRow.SheetName, charRow.RowNum, col, []);
                }
            }
        }
    }
}
