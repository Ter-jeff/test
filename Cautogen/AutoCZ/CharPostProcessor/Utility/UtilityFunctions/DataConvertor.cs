using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.AC;
using Cautogen.AutoCZ.CharPostProcessor.Controller;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;
using Cautogen.common.IgxlDataExtension;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;
using IgxlLib.IgxlSheets.MultiSheet.MultiTimeSet;

using LogLib.Utility;

using HardIpReference = Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.HardIpReference;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions
{
    public class DataConvertor
    {
        public static Dictionary<string, string> SpecialHacInstanceRows = new Dictionary<string, string>();

        private static readonly Dictionary<string, string> _dcCategoryHlnvDict = new Dictionary<string, string>();

        public static string GetGenAcSpec(string timeSetVersion, AcSpecSheet acSpecSheet, string testInstanceName)
        {
            try
            {
                if (acSpecSheet == null)
                {
                    return "";
                }

                string acSpecName = AcTSetCategoryMapSingleton.Instance().GetCategory(timeSetVersion);
                if (acSpecName.ToUpper() == "TBD")
                {
                    if (!File.Exists(Path.Combine(LocalSpecs.TimeSetFolder, timeSetVersion + ".txt")))
                    {
                        return "TBD";
                    }
                    ComTimeSetBasicSheet comTimeSetBasicSheet = MultiTimeSetSheetReader
                        .ReadTimeSetTxt1P4(new List<string>() { Path.Combine(LocalSpecs.TimeSetFolder, timeSetVersion + ".txt") })
                        .TimeSetBasicSheetsList.FirstOrDefault();
                    if (comTimeSetBasicSheet == null)
                    {
                        return "TBD";
                    }

                    string category = ComTimeSetBasicSheet.GetTimeSetCategory(timeSetVersion);
                    string acName = acSpecSheet.AcCatalogContainsTimeSet(comTimeSetBasicSheet, category);
                    if (!string.IsNullOrEmpty(acName))
                    {
                        return acName;
                    }

                    acName = acSpecSheet.AcCatalogContainsTimeSet(comTimeSetBasicSheet);
                    if (!string.IsNullOrEmpty(acName))
                    {
                        return acName;
                    }

                    string newAcSpecName = "AC_" + timeSetVersion;
                    if (!newAcSpecName.ToUpper().StartsWith("TBD") &&
                        !acSpecSheet.CategoryList.Exists(x => x.Equals(newAcSpecName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        acSpecSheet.CategoryList.Add(newAcSpecName);
                        foreach (ComTimeSetBasicSheet.TSetEqnVarMap tset in comTimeSetBasicSheet.AllTSetEqnVariable)
                        {
                            foreach (KeyValuePair<string, double> variable in tset.DictVariable)
                            {
                                string symbol = variable.Key;
                                foreach (AcSpec acSpec in acSpecSheet.Rows)
                                {
                                    if (string.IsNullOrEmpty(acSpec.Symbol))
                                    {
                                        continue;
                                    }

                                    if (acSpec.Symbol.Equals(symbol, StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        string value = variable.Key;
                                        var item = new CategoryInSpec(newAcSpecName, value, value, value);
                                        acSpec.AddCategory(item);
                                    }
                                    else if (acSpec.Symbol.Equals("Using TSet", StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        var item = new CategoryInSpec(newAcSpecName, timeSetVersion, "", "");
                                        acSpec.AddCategory(item);
                                    }
                                    else
                                    {
                                        if (acSpec.CategoryList.Exists(x => x.Name.Equals(acSpecName, StringComparison.CurrentCultureIgnoreCase)))
                                        {
                                            CategoryInSpec row = acSpec.CategoryList.Find(x => x.Name.Equals(acSpecName, StringComparison.CurrentCultureIgnoreCase));
                                            var item = new CategoryInSpec(newAcSpecName, row.Typ, row.Min, row.Max);
                                            acSpec.AddCategory(item);
                                        }
                                        else
                                        {
                                            CategoryInSpec row = acSpec.CategoryList.First();
                                            var item = new CategoryInSpec(newAcSpecName, row.Typ, row.Min, row.Max);
                                            acSpec.AddCategory(item);

                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (FormatException)
            {
                throw new Exception("");
            }
            return "";
        }

        public static InstanceRow ConvertPlanItemToInstanceRow(CharPlanItem planItem, string functionName, bool genPmode = true, bool useNewTChar = false, bool isCSharp = false)
        {
            InstanceRow instRow = planItem.IsUseRtosCmd ? SearchRtosInstanceRow(planItem) : null;

            if (instRow != null)
            {
                return instRow;
            }

            VbtFunction vbtFunction = null;

            if (!isCSharp)
            {
                vbtFunction = BasMain.VbtFunctionLib.GetFunctionByName(functionName);
                vbtFunction.SetParamValue("Interpose_Prepat", ConvertCharCondition(planItem.CharCondition));

            }
            else
            {
                vbtFunction = BasMain.VbtFunctionLib.GetFunctionByName("FuncTestCharMain");
                vbtFunction.SetParamValue("interposePrePat", ConvertCharCondition(planItem.CharCondition));
            }

            if (useNewTChar)
            {
                _SetTCharParameters(planItem, vbtFunction, isCSharp);
                _SetSelSramParameters(planItem, vbtFunction, genPmode, isCSharp);
            }
            else
            {
                _SetVbtParameters(planItem, vbtFunction, false);
                _SetSelSramParameters(planItem, vbtFunction, genPmode);
            }

            InstanceRow instanceRow = null;

            if (!isCSharp)
            {
                instanceRow = new InstanceRow
                {
                    ColumnA = planItem.SheetName + ",Row_" + planItem.RowNum,
                    VbtName = vbtFunction.FunctionName,
                    VbtType = "VBT",
                    ArgList = vbtFunction.Parameters,
                    DcCategory = _GetCategoryFromTestName(planItem.IsOverWriteVoltage,
                    planItem.TestInstanceName,
                    planItem.DcCategory,
                    false),
                    DcSelector = planItem.DcSelector,
                    AcCategory = string.IsNullOrEmpty(planItem.AcCategory) ?
                    GetGenAcSpec(planItem.Timeset, LocalSpecs.TestProgram.AcSpecSheets.FirstOrDefault(), planItem.TestInstanceName)
                    : planItem.AcCategory,
                    AcSelector = planItem.AcSelector,
                    PinLevels = planItem.Levels,
                    TimeSets = planItem.Timeset,
                    Args = vbtFunction.ArgList.ConvertAll(a => a)  // shallow copy
                };
            }
            else
            {
                //CSharp
                instanceRow = new InstanceRow
                {
                    ColumnA = planItem.SheetName + ",Row_" + planItem.RowNum,
                    VbtName = vbtFunction.FullFunctionName,
                    VbtType = ".NET",
                    ArgList = vbtFunction.Parameters,
                    DcCategory = _GetCategoryFromTestName(planItem.IsOverWriteVoltage,
                    planItem.TestInstanceName,
                    planItem.DcCategory,
                    false),
                    DcSelector = planItem.DcSelector,
                    AcCategory = string.IsNullOrEmpty(planItem.AcCategory) ?
                    GetGenAcSpec(planItem.Timeset, LocalSpecs.TestProgram.AcSpecSheets.FirstOrDefault(), planItem.TestInstanceName)
                    : planItem.AcCategory,
                    AcSelector = planItem.AcSelector,
                    PinLevels = planItem.Levels,
                    TimeSets = planItem.Timeset,
                    Args = vbtFunction.ArgList.ConvertAll(a => a)  // shallow copy
                };
            }


            return instanceRow;
        }

        public static InstanceRow ConvertPlanItemToInstanceRow(List<CharPlanItem> planItems, string functionName, bool genPmode = true, bool useNewTChar = false, bool isCSharp = false)
        {
            //var instRow = (planItems.Items..IsUseRtosCmd) ? SearchRtosInstanceRow(planItem) : null;
            InstanceRow instanceRow = null;
            CharPlanItem topDieItem = planItems.FirstOrDefault(i => i.Die == "TOP");
            CharPlanItem botDieItem = planItems.FirstOrDefault(i => i.Die == "BOT");
            VbtFunction vbtFunction = null;
            vbtFunction = BasMain.VbtFunctionLib.GetFunctionByName("FuncTestCharMain");

            vbtFunction.SetParamValue("interposePrePat", ConvertCharCondition(topDieItem.CharCondition));

            _SetTCharParameters(planItems, vbtFunction, isCSharp);


            instanceRow = new InstanceRow
            {
                ColumnA = topDieItem.SheetName + ",Row_" + topDieItem.RowNum,
                VbtName = vbtFunction.FullFunctionName,
                VbtType = ".NET",
                ArgList = vbtFunction.Parameters,
                DcCategory = _GetCategoryFromTestName(topDieItem.IsOverWriteVoltage,
                topDieItem.TestInstanceName,
                topDieItem.DcCategory,
                false),
                DcSelector = topDieItem.DcSelector,
                AcCategory = string.IsNullOrEmpty(topDieItem.AcCategory) ?
                GetGenAcSpec(topDieItem.Timeset, LocalSpecs.TestProgram.AcSpecSheets.FirstOrDefault(), topDieItem.TestInstanceName)
                : topDieItem.AcCategory,
                AcSelector = topDieItem.AcSelector,
                PinLevels = topDieItem.Levels,
                TimeSets = topDieItem.Timeset,
                Args = vbtFunction.ArgList.ConvertAll(a => a)  // shallow copy
            };
            return instanceRow;

        }

        public static InstanceRow ConvertProgInstanceToInstanceRow(CharPlanItem planItem, InstanceRow progInstance, bool genCSharp = false)
        {

            var instRow = new InstanceRow
            {
                VbtName = progInstance.VbtName,
                VbtType = progInstance.VbtType,
                DcCategory = planItem.IsOverWriteVoltage ? planItem.DcCategory : DetProdOrCharInfo(progInstance.DcCategory, planItem.DcCategory),
                DcSelector = DetProdOrCharInfo(progInstance.DcSelector, planItem.DcSelector),
                AcCategory = DetProdOrCharInfo(progInstance.AcCategory, planItem.AcCategory),
                AcSelector = DetProdOrCharInfo(progInstance.AcSelector, planItem.AcSelector),
                PinLevels = DetProdOrCharInfo(progInstance.PinLevels, planItem.Levels),
                TimeSets = progInstance.TimeSets
            };

            var charArgs = progInstance.Args.ToList();
            VbtFunction vbtFunc = genCSharp ? BasMain.VbtFunctionLib.GetFunctionByName(instRow.VbtName.Split('.').LastOrDefault()) : BasMain.VbtFunctionLib.GetFunctionByName(instRow.VbtName);
            vbtFunc.ArgList = charArgs;

            ApplyForceConditionInterposePrePat(planItem, vbtFunc, charArgs);

            //Support for C# module
            if (vbtFunc.FunctionName.Equals("IDSCurrent", StringComparison.OrdinalIgnoreCase) && genCSharp)
            {
                ApplyIdsCurrentCsharpArgs(planItem, vbtFunc, charArgs);
            }

            // update arguments for ids vbt "IDS_main_current"
            if (vbtFunc.FunctionName.Equals("IDS_main_current", StringComparison.OrdinalIgnoreCase) && !genCSharp)
            {
                ApplyIdsMainCurrentArgs(planItem, vbtFunc, charArgs);
            }

            //Get UserDef6
            string userdef6 = GetInfoFromTestName(planItem.TestInstanceName, 7);
            if (!userdef6.Equals("X", StringComparison.OrdinalIgnoreCase))
            {
                vbtFunc.SetParamValue("ForceCMD", userdef6.Replace("nbsp", " "));
            }

            //Get UserDef9 for HAC digsrc
            if (planItem.MeasType == "HAC" && planItem.TestInstanceName.ToUpper().Contains("RING"))
            {
                string userdef9 = GetInfoFromTestName(planItem.TestInstanceName, 10);
                _SetHardIpDigSrc(userdef9, vbtFunc);
            }

            _SetVbtParameters(planItem, vbtFunc, true);

            if (instRow.TimeSets.ToUpper() == "NA")
            {
                instRow.TimeSets = planItem.Timeset;
            }

            instRow.EdgeSets = progInstance.EdgeSets;
            instRow.MixedSignalTiming = progInstance.MixedSignalTiming;
            instRow.ArgList = vbtFunc.Parameters;
            instRow.Args = new List<string>();

            foreach (string arg in vbtFunc.ArgList)
            {
                instRow.Args.Add(arg);
            }
            instRow.Overlay = progInstance.Overlay;
            instRow.CalledAs = progInstance.CalledAs;

            return instRow;
        }

        private static void ApplyForceConditionInterposePrePat(CharPlanItem planItem, VbtFunction vbtFunc, List<string> charArgs)
        {
            var forceConditionName = new List<string> { "CZ_PrePatt", "Interpose_Prepat", "InterFunc_PrePat", "interposePrePat", "interposePreMeasure" };

            // update vbt argument "Interpose_Prepat"
            string forceItem = forceConditionName.FirstOrDefault(p => vbtFunc.Parameters.Split(',').Contains(p, StringComparer.OrdinalIgnoreCase));
            if (forceItem == null || planItem.ExtendInit)
            {
                return;
            }

            int index = vbtFunc.Parameters.Split(',').ToList().FindIndex(s => s.Equals(forceItem, StringComparison.OrdinalIgnoreCase));
            if (charArgs.Count <= index)
            {
                return;
            }

            //Replace shmoo sweep pin condition to shmoo global
            if (planItem.CharShmooSetup != null)
            {
                ReplaceShmooSweepPinCondition(planItem, charArgs, index);
            }

            vbtFunc.ArgList[index] = _GetNewCharInputStr(charArgs[index], ConvertCharCondition(planItem.CharCondition));
        }

        private static void ReplaceShmooSweepPinCondition(CharPlanItem planItem, List<string> charArgs, int index)
        {
            IGLinkProcessor.DataStructure.ShmooData.ShmooPin shmooSetup = planItem.CharShmooSetup.ShmooPins.FirstOrDefault();
            string shmooSweepPin = shmooSetup.SweepPinName;
            string[] originalInterposePrePatList = charArgs[index].Split(';');

            if (LocalSpecs.ProgInfo.PinDic.TryGetValue(shmooSweepPin, out string pin))
            {
                if (Regex.IsMatch(shmooSetup.SweepType, @"Global\s*Spec", RegexOptions.IgnoreCase))
                {
                    for (int i = 0; i < originalInterposePrePatList.Length; ++i)
                    {
                        string item = originalInterposePrePatList[i];
                        string originalInterposePrePatPin = item.Split(':').First();
                        if (pin == originalInterposePrePatPin)
                        {
                            originalInterposePrePatList[i] = $"{pin}:V:Shmoo_Glb";
                        }
                    }
                }
            }
            charArgs[index] = string.Join(";", originalInterposePrePatList);
        }

        private static void ApplyIdsCurrentCsharpArgs(CharPlanItem planItem, VbtFunction vbtFunc, List<string> charArgs)
        {
            IEnumerable<string> idsPinsFromPlan = planItem.TestInstanceName.Split(',').Where(x => x.Split('_').Length >= 6).Select(x => x.Split('_')[5]);
            IEnumerable<string> idsDcvsPins = LocalSpecs.ProgInfo.PinTypeInChannelDic.Where(x => idsPinsFromPlan.Any(y => y == x.Key.Replace("_", "")) && x.Value.StartsWith("DCVS")).Select(x => x.Key);
            IEnumerable<string> idsDcviPins = LocalSpecs.ProgInfo.PinTypeInChannelDic.Where(x => idsPinsFromPlan.Any(y => y == x.Key.Replace("_", "")) && x.Value == "DCVI").Select(x => x.Key);
            IEnumerable<string> autoRangePins = new List<string>().Concat(idsDcvsPins).Concat(idsDcviPins);

            int measuredPinsIndex = vbtFunc.Parameters.Split(',').ToList().FindIndex(s => s.Equals("measuredPins", StringComparison.OrdinalIgnoreCase));

            if (measuredPinsIndex != -1 && charArgs.Count > measuredPinsIndex)
            {
                _ = idsDcvsPins.Any() ? string.Join(",", idsDcvsPins) : "";
                _ = idsDcviPins.Any() ? string.Join(",", idsDcviPins) : "";


                vbtFunc.ArgList[measuredPinsIndex] = idsDcvsPins.Any() ? string.Join(",", idsDcvsPins) : "";
            }

            int autoRangePinIndex = vbtFunc.Parameters.Split(',').ToList().FindIndex(s => s.Equals("autoRangePins", StringComparison.OrdinalIgnoreCase));
            if (autoRangePinIndex != -1 && charArgs.Count > autoRangePinIndex)
            {
                vbtFunc.ArgList[autoRangePinIndex] = autoRangePins.Any() ? string.Join(",", autoRangePins) : "";
            }

            // argument "initialCRFromLimit"
            int flowLimitForInitIRangeIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("initialCRFromLimit", StringComparison.OrdinalIgnoreCase));
            if (flowLimitForInitIRangeIndex != -1 && charArgs.Count > flowLimitForInitIRangeIndex)
            {
                vbtFunc.ArgList[flowLimitForInitIRangeIndex] = "False";
            }
            // argument "fuseEnable"
            int fuseEnableIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("fuseEnable", StringComparison.OrdinalIgnoreCase));
            if (fuseEnableIndex != -1 && charArgs.Count > fuseEnableIndex)
            {
                vbtFunc.ArgList[fuseEnableIndex] = "False";
            }
            // argument "dictName"
            int dictNameIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("dictName", StringComparison.OrdinalIgnoreCase));
            if (dictNameIndex != -1 && charArgs.Count > dictNameIndex)
            {
                vbtFunc.ArgList[dictNameIndex] = "";
            }
        }

        private static void ApplyIdsMainCurrentArgs(CharPlanItem planItem, VbtFunction vbtFunc, List<string> charArgs)
        {
            IEnumerable<string> idsPinsFromPlan = planItem.TestInstanceName.Split(',').Where(x => x.Split('_').Length >= 6).Select(x => x.Split('_')[5]);
            IEnumerable<string> idsDcvsPins = LocalSpecs.ProgInfo.PinTypeInChannelDic.Where(x => idsPinsFromPlan.Any(y => y == x.Key.Replace("_", "")) && x.Value.StartsWith("DCVS")).Select(x => x.Key);
            IEnumerable<string> idsDcviPins = LocalSpecs.ProgInfo.PinTypeInChannelDic.Where(x => idsPinsFromPlan.Any(y => y == x.Key.Replace("_", "")) && x.Value == "DCVI").Select(x => x.Key);
            IEnumerable<string> autoRangePins = new List<string>().Concat(idsDcvsPins).Concat(idsDcviPins);

            // argument "DCVS_Power_Pin"
            int dcvsPowerPinIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("DCVS_Power_Pin", StringComparison.OrdinalIgnoreCase));
            if (dcvsPowerPinIndex != -1 && charArgs.Count > dcvsPowerPinIndex)
            {
                vbtFunc.ArgList[dcvsPowerPinIndex] = idsDcvsPins.Any() ? string.Join(",", idsDcvsPins) : "";
            }
            // argument "DCVI_Power_Pin"
            int dcviPowerPinIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("DCVI_Power_Pin", StringComparison.OrdinalIgnoreCase));
            if (dcviPowerPinIndex != -1 && charArgs.Count > dcviPowerPinIndex)
            {
                vbtFunc.ArgList[dcviPowerPinIndex] = idsDcviPins.Any() ? string.Join(",", idsDcviPins) : "";
            }
            // argument "DCVS_OtherPower_Pin"
            int dcvsOtherPowerPinIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("DCVS_OtherPower_Pin", StringComparison.OrdinalIgnoreCase));
            if (dcvsOtherPowerPinIndex != -1 && charArgs.Count > dcvsOtherPowerPinIndex)
            {
                vbtFunc.ArgList[dcvsOtherPowerPinIndex] = "";
            }
            // argument "FlowLimitForInitIRange"
            int flowLimitForInitIRangeIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("FlowLimitForInitIRange", StringComparison.OrdinalIgnoreCase));
            if (flowLimitForInitIRangeIndex != -1 && charArgs.Count > flowLimitForInitIRangeIndex)
            {
                vbtFunc.ArgList[flowLimitForInitIRangeIndex] = "False";
            }
            // argument "FlowLimitForInitIRange"
            int fuseEnableIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("Fuse_Enable", StringComparison.OrdinalIgnoreCase));
            if (fuseEnableIndex != -1 && charArgs.Count > fuseEnableIndex)
            {
                vbtFunc.ArgList[fuseEnableIndex] = "False";
            }
            // argument "AutoRange_Pin"
            int autoRangePinIndex = vbtFunc.Parameters.Split(',').ToList()
                .FindIndex(s => s.Equals("AutoRange_Pin", StringComparison.OrdinalIgnoreCase));
            if (autoRangePinIndex != -1 && charArgs.Count > autoRangePinIndex)
            {
                vbtFunc.ArgList[autoRangePinIndex] = autoRangePins.Any() ? string.Join(",", autoRangePins) : "";
            }
        }

        private static void _SetHardIpDigSrc(string userdef9, VbtFunction vbtFunction)
        {
            Match userdef9Ring = Regex.Match(userdef9, @"(?<Name>[A-Za-z]+)(?<Bits>\d+)");
            if (!userdef9Ring.Success)
            {
                return;
            }

            int index =
                vbtFunction.Parameters.Split(',').ToList().FindIndex(
                    s => s.Equals("DigSrc_Assignment", StringComparison.OrdinalIgnoreCase));
            if (index == -1)
            {
                return;
            }

            string digSrcAssign = vbtFunction.ArgList[index];
            List<string> digSrcAssignList = digSrcAssign.Split(';').ToList();

            string name = userdef9Ring.Groups["Name"].Value;
            string bits = userdef9Ring.Groups["Bits"].Value;
            for (int i = 0; i < digSrcAssignList.Count; i++)
            {
                string assign = digSrcAssignList[i];
                if (assign.Split('=').Length < 2)
                {
                    continue;
                }

                string argName = assign.Split('=')[0];
                string argBits = assign.Split('=')[1];
                if (!string.IsNullOrEmpty(argName) && !string.IsNullOrEmpty(argBits))
                {
                    string printBits = bits.PadLeft(argBits.Length, '0');
                    digSrcAssignList[i] = argName + "=" + printBits;
                }
            }
            vbtFunction.SetParamValue("DigSrc_Assignment", string.Join(";", digSrcAssignList));
        }

        public static InstanceRow SearchRtosInstanceRow(CharPlanItem planItem)
        {
            List<InstanceRow> instList = ProdProg.AllTestInstances.FindAll(p => p.TestName.ToLower().Contains(planItem.Payload1.ToLower()));
            if (instList.Count <= 0)
            {
                return null;
            }

            InstanceRow tpInst = instList.FirstOrDefault(p => p.VbtName.Contains(planItem.Voltage) && p.VbtName.Contains(planItem.SheetName)) ??
                         instList.FirstOrDefault(p => p.VbtName.Equals("RTOS_RunScenario_CZ", StringComparison.OrdinalIgnoreCase)) ??
                         instList.FirstOrDefault(p => p.VbtName.Equals("RTOS_RunScenario", StringComparison.OrdinalIgnoreCase)) ??
                         instList[0];

            VbtFunction function = BasMain.VbtFunctionLib.GetFunctionByName(tpInst.VbtName + "_CZ");
            if (function.FileName == null)
            {
                function = BasMain.VbtFunctionLib.GetFunctionByName(tpInst.VbtName);
            }

            // fill args
            string[] argNames = tpInst.ArgList.Split(',');
            if (argNames.Length > 1)
            {
                for (int i = 0; i < argNames.Length; i++)
                {
                    if (i < tpInst.Args.Count)
                    {
                        function.SetParamValue(argNames[i], tpInst.Args[i]);
                    }
                }
            }
            else
            {
                int i = 0;
                while (i < function.ArgList.Count && i < tpInst.Args.Count)
                {
                    function.ArgList[i] = tpInst.Args[i];
                    i++;
                }
            }

            function.SetParamValue("Interpose_Prepat", ConvertCharCondition(planItem.CharCondition));
            string userdef6 = GetInfoFromTestName(planItem.TestInstanceName, 7);//Get UserDef6 
            if (!userdef6.Equals("X", StringComparison.OrdinalIgnoreCase))
            {
                function.SetParamValue("ForceCMD", userdef6.Replace("nbsp", " "));
            }

            return new InstanceRow
            {
                VbtName = function.FunctionName,
                VbtType = "VBT",
                ArgList = function.Parameters,
                Args = function.ArgList.ToList(),  // deep copy
                DcCategory = planItem.IsOverWriteVoltage ? planItem.DcCategory : tpInst.DcCategory,
                DcSelector = tpInst.DcSelector,
                AcCategory = tpInst.AcCategory,
                AcSelector = tpInst.AcSelector,
                PinLevels = tpInst.PinLevels,
                TimeSets = tpInst.TimeSets
            };
        }
        public static string ConvertPatternIndex(Dictionary<string, string> convertDict, string value)
        {
            IEnumerable<string> resultList = value.Split(',').Select(x =>
            {
                foreach (KeyValuePair<string, string> convert in convertDict)
                {
                    if (x.Split(':')[0].Equals(convert.Key))
                    {
                        x = x.Replace(convert.Key, convert.Value);
                        break;
                    }
                }
                return x;
            });
            return string.Join(",", resultList);
        }

        public static string ConvertCSharpPatternIndex(Dictionary<string, string> convertDict, string value)
        {
            //vbtFunction.SetParamValue("digSrcPin", planItem.DigSrcPin.Split(':')[1].Replace("|", ""));
            //vbtFunction.SetParamValue("digSrcEquation", ConvertCSharpPatternIndex(planItem.ArgPatternIndexConvertedDsscDict, planItem.DigSrcEQ));
            string[] patternIndexList = new string[convertDict.Count];
            string[] sourceSegList = value.Split(',');

            for (int i = 0; i < patternIndexList.Length; ++i)
            {
                patternIndexList[i] = "|";
            }

            int sgmtIdx = 0;

            foreach (string sourceSeg in sourceSegList)
            {
                string splitSegBit = convertDict[sourceSeg.Split(':')[0]].Split(':')[0].Replace("INIT", "").Replace("PL", "");
                int bitIdx = -1;
                int.TryParse(splitSegBit, out bitIdx);

                patternIndexList[bitIdx - 1] = $"sgmt{sgmtIdx}";
                sgmtIdx += 1;

            }

            return string.Join("", patternIndexList);
        }



        public static string ConvertPwrPatternIndex(Dictionary<string, string> convertDict, string value)
        {
            IEnumerable<string> resultList = value.Split('_').Select(x =>
            {
                foreach (KeyValuePair<string, string> convert in convertDict)
                {
                    if (x.Equals(convert.Key))
                    {
                        x = x.Replace(convert.Key, convert.Value);
                        break;
                    }
                }
                return x;
            });
            return string.Join("_", resultList);
        }
        private static string _ExtendPowerRunScenarioIdex(Dictionary<string, List<string>> convertDict, string value)
        {
            var newPws = new List<string>();
            var pwsDict = new Dictionary<string, string>();
            int loopSegIdx = 0;
            string temp = "";
            foreach (string segment in value.Split('_'))
            {
                loopSegIdx++;
                if (loopSegIdx % 2 == 1)
                {
                    temp = segment;
                    continue;
                }
                pwsDict[temp] = segment;
            }
            foreach (KeyValuePair<string, string> pwr in pwsDict)
            {
                if (!Regex.IsMatch(pwr.Key, @"^(INIT|PL)\d+$", RegexOptions.IgnoreCase))
                {
                    newPws.Add(string.Format($"{pwr.Key}_{pwr.Value}"));
                    continue;
                }
                else
                {
                    string target = convertDict.Keys.FirstOrDefault(x => string.Equals(x, pwr.Key, StringComparison.OrdinalIgnoreCase));
                    if (!string.IsNullOrEmpty(target))
                    {
                        foreach (string newName in convertDict[target])
                        {
                            newPws.Add(string.Format($"{newName}_{pwr.Value}"));
                        }
                    }
                }
            }
            return string.Join("_", newPws);
        }
        private static void _SetTCharParameters(CharPlanItem planItem, VbtFunction vbtFunction, bool isCSharp = false)
        {
            if (!isCSharp)
            {
                vbtFunction.SetParamValue("INIT_PATSET", planItem.TestInstanceName + "INIT");
                vbtFunction.SetParamValue("PL_PATSET", planItem.TestInstanceName + "PL");
                vbtFunction.SetParamValue("BypassShmooHole", planItem.BypassShmooHole);
                vbtFunction.SetParamValue("Char_Setup_Name", planItem.CharShmooSetup.ShmooSetupName);
                vbtFunction.SetParamValue("Harvest_Header", planItem.HarvFstp);
                vbtFunction.SetParamValue("Power_Run_Scenario", _ExtendPowerRunScenarioIdex(planItem.ArgPatternIndexConvertedPrsDict, planItem.PowerRunScenario));
                vbtFunction.SetParamValue("Wait", ConvertPatternIndex(planItem.ArgPatternIndexConvertedRetDict, ConvertWaitFromOldtoNew(planItem.Wait)));
                vbtFunction.SetParamValue("Ret_Ramp_Setting", ConvertPatternIndex(planItem.ArgPatternIndexConvertedRetDict, planItem.RetentionRamp));
                vbtFunction.SetParamValue("DigSrc_BitSize", ConvertPatternIndex(planItem.ArgPatternIndexConvertedDsscDict, planItem.DigSrcBitSize));
                vbtFunction.SetParamValue("DigSrc_Seg", ConvertPatternIndex(planItem.ArgPatternIndexConvertedDsscDict, planItem.DigSrcSeg));
                vbtFunction.SetParamValue("DigSrc_DigSrcPin", ConvertPatternIndex(planItem.ArgPatternIndexConvertedDsscDict, planItem.DigSrcPin));
                vbtFunction.SetParamValue("digSrc_EQ", ConvertPatternIndex(planItem.ArgPatternIndexConvertedDsscDict, planItem.DigSrcEq));
                vbtFunction.SetParamValue("Order_LSB", planItem.IsDateNeedReverse ? "TRUE" : "FALSE");
                vbtFunction.SetParamValue("One_Time_INIT", planItem.OneTimeInit);
                vbtFunction.SetParamValue("ApplyVoltageFromBinCut", planItem.ApplyVoltageFromBinCut);
                vbtFunction.SetParamValue("UserFunction", planItem.UserFunction);
                vbtFunction.SetParamValue("HarvestPinGrpOtherFail", planItem.HarvestPinGrpOtherFail);
                vbtFunction.SetParamValue("EnableCoreHarvest", planItem.EnableCoreHarvest);
                vbtFunction.SetParamValue("EnableCoreMask", planItem.EnableCoreMask);
                vbtFunction.SetParamValue("PinGrp_SpecifyMask", planItem.PinGrpSpecifyMask);
                vbtFunction.SetParamValue("SSN_SpecifyMask", planItem.SsnSpecifyMask);

            }
            else
            {
                vbtFunction.SetParamValue("initPatset", planItem.TestInstanceName + "INIT");
                vbtFunction.SetParamValue("plpatset", planItem.TestInstanceName + "PL");
                vbtFunction.SetParamValue("bypassShmooHole", planItem.BypassShmooHole);
                //vbtFunction.SetParamValue("Harvest_Header", planItem.Harv_FSTP);
                vbtFunction.SetParamValue("powerRunScenario", _ExtendPowerRunScenarioIdex(planItem.ArgPatternIndexConvertedPrsDict, planItem.PowerRunScenario));
                vbtFunction.SetParamValue("wait", ConvertPatternIndex(planItem.ArgPatternIndexConvertedRetDict, ConvertWaitFromOldtoNew(planItem.Wait)));
                vbtFunction.SetParamValue("retRampSetting", ConvertPatternIndex(planItem.ArgPatternIndexConvertedRetDict, planItem.RetentionRamp));

                ConvertCSharpPatternDigSrc(vbtFunction, planItem);

                vbtFunction.SetParamValue("orderMSB", planItem.IsDateNeedReverse ? "TRUE" : "FALSE");
                vbtFunction.SetParamValue("oneTimeInit", planItem.OneTimeInit);
                vbtFunction.SetParamValue("ateTestCondition", planItem.ApplyVoltageFromBinCut);
                //vbtFunction.SetParamValue("userFunction", planItem.UserFunction);
            }

        }
        private static void _SetTCharParameters(List<CharPlanItem> planItems, VbtFunction vbtFunction, bool isCSharp = false)
        {

            CharPlanItem topDieItem = planItems.FirstOrDefault(i => i.Die == "TOP");
            CharPlanItem botDieItem = planItems.FirstOrDefault(i => i.Die == "BOT");

            string mtdInitPatSetCombine = $"{topDieItem.TestInstanceName}_{topDieItem.Die}_INIT+{botDieItem.TestInstanceName}_{botDieItem.Die}_INIT";
            string mtdPlPatSetCombine = $"{topDieItem.TestInstanceName}_{topDieItem.Die}_PL+{botDieItem.TestInstanceName}_{botDieItem.Die}_PL";

            vbtFunction.SetParamValue("initPatset", mtdInitPatSetCombine);
            vbtFunction.SetParamValue("plpatset", mtdPlPatSetCombine);
            vbtFunction.SetParamValue("bypassShmooHole", topDieItem.BypassShmooHole);
            //vbtFunction.SetParamValue("Harvest_Header", planItem.Harv_FSTP);
            vbtFunction.SetParamValue("powerRunScenario", _ExtendPowerRunScenarioIdex(topDieItem.ArgPatternIndexConvertedPrsDict, topDieItem.PowerRunScenario));
            vbtFunction.SetParamValue("wait", ConvertPatternIndex(topDieItem.ArgPatternIndexConvertedRetDict, ConvertWaitFromOldtoNew(topDieItem.Wait)));
            vbtFunction.SetParamValue("retRampSetting", ConvertPatternIndex(topDieItem.ArgPatternIndexConvertedRetDict, topDieItem.RetentionRamp));

            //ConvertCSharpPatternDigSrc(vbtFunction, planItem);

            vbtFunction.SetParamValue("orderMSB", topDieItem.IsDateNeedReverse ? "TRUE" : "FALSE");
            vbtFunction.SetParamValue("oneTimeInit", topDieItem.OneTimeInit);
            vbtFunction.SetParamValue("ateTestCondition", topDieItem.ApplyVoltageFromBinCut);
        }

        private static void ConvertCSharpPatternDigSrc(VbtFunction vbtFunction, CharPlanItem planItem)
        {
            string digSrcEquation = "";
            var digSrcAssignment = new List<string>();

            //To search digital source pattern info
            HardIpReference patInfo = null;
            if (!string.IsNullOrEmpty(planItem.DigSrc))
            {
                foreach (string srcStr in planItem.DigSrc.Split(','))
                {
                    if (!string.IsNullOrEmpty(srcStr))
                    {
                        string patternName = srcStr.Split(':')[0];
                        string srcType = srcStr.Split(':')[1];
                        string binStr = srcStr.Split(':')[2];//Defined in CharPlan Userdef6

                        if (srcType.Equals("DigSrc", StringComparison.OrdinalIgnoreCase))
                        {
                            patInfo = SearchInfo.GetHardIpInfo(patternName);
                            break;
                        }
                    }


                }
            }


            string[] patternIndexList = new string[(planItem.AllPatternsDict.Count * 2) - 1];
            if (string.IsNullOrEmpty(planItem.DigSrcPin))
            {
                return;
            }

            string sourcePin = planItem.DigSrcPin.Split(',').ToList().Select(p => p.Split(':')[1]).ToList().Distinct().FirstOrDefault().ToString();
            string[] sourceEqList = planItem.DigSrcEq.Split(',');
            string[] sourceSegList = planItem.DigSrcSeg.Split(',');
            string[] sourceAssign = planItem.DigSrcAssignment.Split(',');

            for (int i = 0; i < patternIndexList.Length; ++i)
            {
                patternIndexList[i] = i % 2 == 1 ? "|" : "";
            }

            int sgmtIdx = 0;


            foreach (string sourceSeg in sourceEqList)
            {
                //This mapping need to re-calculate the index for INIT1 ... INIT2 ... INITn ...
                //It will re-arrange init/payload pattern sequence
                string splitSegBit = planItem.ArgPatternIndexConvertedDsscDict[sourceSeg.Split(':')[0]].Split(':')[0].Replace("INIT", "").Replace("PL", "");
                int bitIdx = -1;
                int.TryParse(splitSegBit, out bitIdx);

                //This pipeline symbol is for C#-Char usage
                patternIndexList[(bitIdx * 2) - 2] = $"sgmt{sgmtIdx}";


                string sgmt = sourceSegList[sgmtIdx].Split('=')[1];

                if (Regex.IsMatch(sgmt, "SELSRAM", RegexOptions.IgnoreCase))
                {
                    var userdefLast = planItem.TestInstanceName.Split('_').Where(x => !string.IsNullOrEmpty(x)).ToList();

                    if (userdefLast.Count > 10)
                    {

                        string selsrmBits = Regex.Replace(userdefLast[10], "selsrm|selsram", "", RegexOptions.IgnoreCase);
                        string selsrmBitsReplace = Regex.Replace(selsrmBits, "s", "", RegexOptions.IgnoreCase);

                        string segment = planItem.DigSrcSeg.Split(':')[1].Split('=')[0];

                        if (!string.IsNullOrEmpty(selsrmBitsReplace))
                        {
                            sgmt = $"sgmt{sgmtIdx}=SELSRAM({string.Join(",", selsrmBits.ToArray())})";
                        }
                        else
                        {
                            sgmt = $"sgmt{sgmtIdx}=SELSRAM()";
                        }
                        digSrcAssignment.Add(sgmt);
                    }
                }
                else
                {
                    //The sourceAssign is come form char plan `allPatterns` count to assign
                    //It's from 0 to 15 the pattern index will follow from 0
                    //Follow original wording (e.g. INIT1,3,4,5) pattern sequence to find pattern digsrc index
                    int.TryParse(sourceSeg.Split(':')[0].Replace("INIT", "").Replace("PL", ""), out int digSrcAssignIdx);
                    digSrcAssignIdx -= 1;
                    sgmt = sourceAssign[digSrcAssignIdx];

                    //Try to convert digsrc assignment
                    try
                    {
                        sgmt = ConvertBinStr(sgmt, patInfo, planItem.IsDateNeedReverse);

                    }
                    catch (Exception)
                    {
                        LogHelper.Error("DigSrc assignment error");
                    }

                    if (string.IsNullOrEmpty(sgmt))
                    {
                        sgmt = ConvertBinStr(sourceAssign[bitIdx - 1], patInfo, planItem.IsDateNeedReverse);

                    }


                    digSrcAssignment.Add($"sgmt{sgmtIdx}={sgmt}");


                }
                sgmtIdx += 1;
            }
            digSrcEquation = string.Join("", patternIndexList);

            vbtFunction.SetParamValue("digSrcPin", sourcePin);
            vbtFunction.SetParamValue("digSrcEquation", digSrcEquation);
            vbtFunction.SetParamValue("digSrcAssignment", string.Join(";", digSrcAssignment));

        }

        private static void _SetVbtParameters(CharPlanItem planItem, VbtFunction vbtFunction, bool useProduction)  // useProduction == true means use the args values in Production instances, else generate it according to patInfo
        {
            vbtFunction.SetParamValue("Init_Patt1", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT1")));
            vbtFunction.SetParamValue("Init_Patt2", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT2")));
            vbtFunction.SetParamValue("Init_Patt3", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT3")));
            vbtFunction.SetParamValue("Init_Patt4", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT4")));
            vbtFunction.SetParamValue("Init_Patt5", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT5")));
            vbtFunction.SetParamValue("Init_Patt6", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT6")));
            vbtFunction.SetParamValue("Init_Patt7", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT7")));
            vbtFunction.SetParamValue("Init_Patt8", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT8")));
            vbtFunction.SetParamValue("Init_Patt9", _ConvertSelsrmPattern(planItem.GetPatternByKey("INIT9")));
            vbtFunction.SetParamValue("Init_Patt10", _ConvertSelsrmPattern(planItem.GetPatternsStartByKey("INIT", 10)));
            vbtFunction.SetParamValue("PayLoad_Patt1", _ConvertSelsrmPattern(planItem.GetPatternByKey("PL1")));
            vbtFunction.SetParamValue("PayLoad_Patt2", _ConvertSelsrmPattern(planItem.GetPatternByKey("PL2")));
            vbtFunction.SetParamValue("PayLoad_Patt3", _ConvertSelsrmPattern(planItem.GetPatternByKey("PL3")));
            vbtFunction.SetParamValue("PayLoad_Patt4", _ConvertSelsrmPattern(planItem.GetPatternByKey("PL4")));
            vbtFunction.SetParamValue("PayLoad_Patt5", _ConvertSelsrmPattern(planItem.GetPatternsStartByKey("PL", 5)));
            vbtFunction.SetParamValue("Power_Run_Scenario", planItem.PowerRunScenario);
            vbtFunction.SetParamValue("Wait", planItem.Wait);
            vbtFunction.SetParamValue("Harvest_Header", planItem.HarvFstp);
            //vbtFunction.SetParamValue("BlockType", planItem.AcCategory.Split('_')[0]);
            vbtFunction.SetParamValue("BypassShmooHole", planItem.BypassShmooHole);
            vbtFunction.SetParamValue("ApplyVoltageFromBinCut", planItem.ApplyVoltageFromBinCut);
            if (planItem.CharShmooSetup != null)
            {
                vbtFunction.SetParamValue("Char_Setup_Name", planItem.CharShmooSetup.ShmooSetupName);
            }

            #region Set Dig Src information
            if (!string.IsNullOrEmpty(planItem.DigSrc))
            {
                var sendBitList = new List<string>();
                var sendBitStrList = new List<string>();
                var sendPinNameList = new List<string>();
                var binStrList = new List<string>();
                foreach (string srcStr in planItem.DigSrc.Split(','))
                {
                    if (srcStr == "")
                    {
                        sendBitList.Add("");
                        sendBitStrList.Add("");
                        sendPinNameList.Add("");
                        binStrList.Add("");
                    }
                    else
                    {
                        if (srcStr.Split(':').Length < 3)
                        {
                            continue;
                        }

                        string patternName = srcStr.Split(':')[0];
                        string srcType = srcStr.Split(':')[1];
                        string binStr = srcStr.Split(':')[2];//Defined in CharPlan Userdef6
                        HardIpReference patInfo = SearchInfo.GetHardIpInfo(patternName);
                        if (patInfo != null)
                        {
                            string sendPinName = patInfo.SendPinName == "" ? "JTAG_TDI" : patInfo.SendPinName;
                            string sendBitStr = patInfo.SendBitStr;
                            string sendBit = patInfo.SendBit.ToString(CultureInfo.InvariantCulture);

                            if (srcType.Equals("DigSrc", StringComparison.OrdinalIgnoreCase))
                            {
                                //if (
                                //    LocalSpecs.SourceMappingItems.Exists(
                                //        p => p.Label.Equals(binStr, StringComparison.OrdinalIgnoreCase)))
                                //    binStr =
                                //        LocalSpecs.SourceMappingItems.FirstOrDefault(
                                //            p => p.Label.Equals(binStr, StringComparison.OrdinalIgnoreCase)).Value;
                                //binStr = ConvertBinStr(binStr, patInfo, planItem.IsDateNeedReverse);
                            }
                            binStr = ConvertBinStr(binStr, patInfo, planItem.IsDateNeedReverse);
                            if (sendBitStr == "" && patInfo.SendBit == 0)
                            {
                                sendBit = binStr.Length.ToString(CultureInfo.InvariantCulture);
                            }

                            sendBitList.Add(sendBit);
                            sendBitStrList.Add(sendBitStr);
                            sendPinNameList.Add(sendPinName);
                        }
                        else if (binStr != "")
                        {
                            sendBitList.Add(binStr.Length.ToString(CultureInfo.InvariantCulture));// bit number
                            sendBitStrList.Add("");
                            sendPinNameList.Add("JTAG_TDI");
                        }
                        else
                        {
                            sendBitList.Add("");
                            sendBitStrList.Add("");
                            sendPinNameList.Add("");
                        }

                        binStrList.Add(binStr);
                    }
                }
                string argBinStr = string.Join(",", binStrList);
                string argSendPinName = string.Join(",", sendPinNameList);
                string argSendBitStr = string.Join(",", sendBitStrList);
                string argSendBit = string.Join(",", sendBitList);


                if (vbtFunction.FunctionName == VbtFunctionLib.FunctionalCharName)
                {
                    vbtFunction.SetParamValue("digsrc_BitSize", argSendBit);
                    vbtFunction.SetParamValue("digsrc_Seg", argSendBitStr);
                    vbtFunction.SetParamValue("digsrc_DigSrcPin", argSendPinName);
                    vbtFunction.SetParamValue("digsrc_eq", "\'" + argBinStr);
                }
                else // Use T-Autogen rule
                {
                    vbtFunction.SetParamValue("DigSrc_Sample_Size", argSendBit);
                    vbtFunction.SetParamValue("DigSrc_Equation", argSendBitStr);
                    vbtFunction.SetParamValue("DigSrc_Pin", argSendPinName);
                    vbtFunction.SetParamValue("DigSrc_Assignment", "\'" + argBinStr);
                }
            }
            #endregion

            if (vbtFunction.FunctionName.ToLower() == VbtFunctionLib.VifName ||
                vbtFunction.FunctionName.ToLower() == VbtFunctionLib.VirName)
            {
                #region Set VIR/VFI common arguments
                HardIpReference patInfo = SearchInfo.GetHardIpInfo(planItem.Payload1);
                if (patInfo == null)
                {
                    return;
                }

                string pinsStatus = GetPinsStatus(patInfo, vbtFunction.FunctionName);
                vbtFunction.SetParamValue("TestLimitPerPin_VIR", pinsStatus.Split(':')[1], false);
                vbtFunction.SetParamValue("TestLimitPerPin_VFI", pinsStatus.Split(':')[1], false);
                if (useProduction)
                {
                    return;
                }

                #region RegisterAssignment
                //DigSrc_Equation: From patInfo file "Send Bit Name"
                vbtFunction.SetParamValue("DigSrc_Equation",
                    patInfo.SendBitName != "" ? patInfo.SendBitName : patInfo.SendBitStr);

                //DigSrc_Sample_Size: Get from "Send Bit" in patInfo file, Like Send Bit: 160  ===> 160
                vbtFunction.SetParamValue("DigSrc_Sample_Size", patInfo.SendBit.ToString("G15"));

                //DigSrc_DataWidth: Get from "Send Bit Str" in patInfo file. Like wdr0_16+wdr1_16+wdr2_16 ===> 16
                vbtFunction.SetParamValue("DigSrc_DataWidth", Regex.Match(patInfo.SendBitStr, @"^[a-zA-Z]+\d+_(?<num>(\d+)).*").Groups["num"].ToString());

                //DigSrc_Pin
                vbtFunction.SetParamValue("DigSrc_Pin", patInfo.SendPinName != "" ? patInfo.SendPinName : "JTAG_TDI");

                #endregion

                #region MeasC

                vbtFunction.SetParamValue("DigCap_Pin", patInfo.CapPinName != "" ? patInfo.CapPinName : "JTAG_TDO");

                //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
                vbtFunction.SetParamValue("DigCap_DataWidth", Regex.Match(patInfo.CapBitStr, @"^[a-zA-Z]+\d+_(?<num>(\d+)).*").Groups["num"].ToString());

                //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
                vbtFunction.SetParamValue("DigCap_Sample_Size", patInfo.CapBit.ToString("G15"));

                #endregion

                vbtFunction.ArgList[0] = planItem.Payload1;

                //Cpu_flag_A
                vbtFunction.SetParamValue("CPUA_Flag_In_Pat", patInfo.MeasSeqStr != "" ? "true" : "false");
                vbtFunction.SetParamValue("TestSequence", patInfo.MeasSeqStr);

                #endregion
            }
            else if (!useProduction &&
                     vbtFunction.FunctionName.Equals(VbtFunctionLib.FunctionalCharName, StringComparison.OrdinalIgnoreCase))
            {
                #region set default value for function_t_char

                vbtFunction.SetParamValue("RelayMode", "1");
                vbtFunction.SetParamValue("WaitFlagA", "-2");
                vbtFunction.SetParamValue("WaitFlagB", "-2");
                vbtFunction.SetParamValue("WaitFlagC", "-2");
                vbtFunction.SetParamValue("WaitFlagD", "-2");
                vbtFunction.SetParamValue("PatternTimeout", "30");

                #endregion
            }
        }

        public static string ConvertWaitFromOldtoNew(string str)
        {
            List<string> strList = str.Split(',').ToList();
            foreach (string split in strList)
            {
                if (split.Trim() == "")
                {
                    continue;
                }
                else
                {
                    if (!decimal.TryParse(split, out _))
                    {
                        return str;
                    }
                }
            }
            int initIdx = 1;
            int plIdx = 1;
            decimal convertNum;
            var resultList = new List<string>();
            if (strList.Count > 5)
            {
                for (int i = 0; i < strList.Count; i++)
                {

                    if (i < 10)
                    {
                        if (decimal.TryParse(strList[i], out convertNum))
                        {
                            resultList.Add(string.Format($"INIT{initIdx}:{convertNum}:"));
                        }
                        initIdx++;
                    }
                    else
                    {
                        if (decimal.TryParse(strList[i], out convertNum))
                        {
                            resultList.Add(string.Format($"PL{plIdx}:{convertNum}:"));
                        }
                        plIdx++;
                    }
                }
            }
            else
            {
                for (int i = 0; i < strList.Count; i++)
                {
                    if (decimal.TryParse(strList[i], out convertNum))
                    {
                        resultList.Add(string.Format($"PL{plIdx}:{convertNum}:"));
                    }
                    plIdx++;
                }
            }
            return string.Join(",", resultList);
        }

        public static string GetMeasType(string instanceName)
        {
            string result = "";
            if (instanceName == "")
            {
                return result;
            }

            string userdefine1 = instanceName.Split('_')[0].ToLower();

            if (Regex.IsMatch(userdefine1, "(dftl|dftlh|mcl|mclh)$"))
            {
                result = "LVCC";
            }
            else if (Regex.IsMatch(userdefine1, "(dfth|mch)$"))
            {
                result = "HVCC";
            }
            else if (Regex.IsMatch(userdefine1, "hf(l|h|lh|hl)$"))
            {
                result = "HF";
            }
            else if (Regex.IsMatch(userdefine1, "(hac|tsmc)$"))
            {
                result = "HAC";
            }
            else if (Regex.IsMatch(userdefine1, "hio$"))
            {
                result = "HIO";
            }
            else
            {
                result = "LVCC";
            }

            return result;
        }

        private static string _GetNewCharInputStr(string orgStr, string newStr)
        {
            var oriList = new CharInputItemList(orgStr);
            var newList = new CharInputItemList(newStr);
            oriList.Extend(newList);
            return oriList.InputStr;
        }

        public static string GetPinsStatus(HardIpReference patInfo, string vbtName = null)
        {
            if (vbtName == null)
            {
                vbtName = patInfo.MeasFStr != "" || patInfo.MeasVPowerPinList.Count > 0 || patInfo.MeasIPowerPinList.Count > 0
                    ? VbtFunctionLib.VifName
                    : VbtFunctionLib.VirName;
            }

            string perPinV = patInfo.MeasVStr.Contains(",") ? "T" : "F";
            string perPinI = patInfo.MeasIStr.Contains(",") ? "T" : "F";
            string perPinF = patInfo.MeasFStr.Contains(",") ? "T" : "F";

            if (Regex.IsMatch(vbtName, "VIR", RegexOptions.IgnoreCase))
            {
                return "VIR:" + perPinV + perPinI + "T";  // MeasR always using per pin use-limit
            }

            return "VFI:" + perPinV + perPinF + perPinI;
        }

        public static bool CheckMixedPins(string pinStatus, string measType)
        {
            string type = measType.Substring(measType.Length - 1, 1).ToUpper();
            if (type == "Z")//MeasType in patInfo or CharPlan Maybe MeasZ, but in vbt parameter only has R
            {
                type = "R";
            }

            string vbtType = pinStatus.Split(':')[0];
            string status = pinStatus.Split(':')[1];
            int index = vbtType.IndexOf(type, StringComparison.Ordinal);
            if (index <= -1)
            {
                return false;
            }

            string result = status.Substring(index, 1);
            return result == "T";
        }

        private static string ConvertBinStr(string binStr, HardIpReference patInfo, bool isReverseData)
        {
            if (Regex.IsMatch(binStr, @"^\d+$"))
            {
                return binStr;
            }

            string[] bitStrArray = patInfo.SendBitStr.Split('+');
            if (!Regex.IsMatch(binStr, @"sgmt\d+"))//EMA remapping
            {
                Cautogen.Utility.EmaMappingItem emaTarget =
                    LocalSpecs.EmaMappingItems.FirstOrDefault(
                        p => p.Pattern.Equals(patInfo.Payload, StringComparison.OrdinalIgnoreCase));
                if (emaTarget != null)
                {
                    string data = emaTarget.GetCaseData(binStr);
                    binStr = data;
                }
            }
            string defaultBinStr = Regex.Match(binStr, @"sgmtdef(?<value>\d+)", RegexOptions.IgnoreCase).Groups["value"].ToString();
            string[] sgmtSets = Regex.Split(binStr, @"(sgmt\d+)", RegexOptions.IgnoreCase);

            var dicBin = new Dictionary<string, string>();

            //sgmtName: sgmt[0-9], sgmtStr: srcData
            for (int i = 0; i < sgmtSets.Length; i++)
            {
                string sgmtName = Regex.Match(sgmtSets[i], @"sgmt\d+", RegexOptions.IgnoreCase).ToString().ToLower();

                if (sgmtName == "" || i + 1 >= sgmtSets.Length)
                {
                    continue;
                }

                string sgmtStr = "";

                if (Regex.IsMatch(sgmtSets[i + 1], "f[01S]+", RegexOptions.IgnoreCase))
                {
                    sgmtStr = sgmtSets[i + 1].Split('f')[1];
                }
                else if (Regex.IsMatch(sgmtSets[i + 1], "g[0-9a-fA-F]+", RegexOptions.IgnoreCase))
                {
                    char[] hexArray =
                        Regex.Match(sgmtSets[i + 1], "g(?<data>[0-9a-fA-F]+)").Groups["data"].ToString().ToCharArray();

                    sgmtStr = (
                        from hex in hexArray
                        select Convert.ToInt32(hex.ToString(CultureInfo.InvariantCulture), 16)
                            into value
                        select Convert.ToString(value, 2)
                                into hexSrc
                        select hexSrc.PadLeft(4, '0')
                        ).Aggregate(sgmtStr, (current, hexSrc) => current + hexSrc);
                }
                dicBin.Add(sgmtName, sgmtStr);
            }

            if (!int.TryParse(defaultBinStr, out int defaultValue))
            {
                defaultValue = 0;
            }

            ConstData.DefaultSigsrcValue = "sgmt_default=" + defaultValue;
            string newBinStr = "";
            if (!string.IsNullOrEmpty(patInfo.SendBitStr))
            {
                foreach (string sgmt in bitStrArray)
                {
                    try
                    {
                        string sgmtName = sgmt.Split('_')[0].ToLower();
                        string sgmtLength = sgmt.Split('_')[1];
                        int bitCount = Convert.ToInt32(sgmtLength);

                        if (dicBin.ContainsKey(sgmtName))
                        {
                            string data = "";
                            if (bitCount > dicBin[sgmtName].Length)
                            {
                                data = dicBin[sgmtName].PadLeft(bitCount, '0');
                            }
                            else if (bitCount < dicBin[sgmtName].Length)
                            {
                                data = dicBin[sgmtName].Substring(dicBin[sgmtName].Length - bitCount, bitCount);
                            }
                            else
                            {
                                data = dicBin[sgmtName];
                            }

                            if (isReverseData)
                            {
                                char[] chararray = data.ToCharArray();
                                Array.Reverse(chararray);
                                data = new string(chararray);
                            }
                            newBinStr += data;
                        }
                        else
                        {
                            string str = Convert.ToString(defaultValue, 2);
                            string newStr = str.PadLeft(bitCount, '0');
                            newBinStr += newStr;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            else
            {
                string dataFromCharPlan = isReverseData
                    ? string.Join("", dicBin.Values.Select(p => new string(p.Reverse().ToArray())))
                    : string.Join("", dicBin.Values.Select(p => p));
                newBinStr = dataFromCharPlan;
            }
            return newBinStr;
        }

        private static string DetProdOrCharInfo(string prodinfo, string charinfo)
        {
            return charinfo != "" ? charinfo : prodinfo;
        }

        public static string GetInfoFromTestName(string tTestName, int index)
        {
            return index < tTestName.Split('_').Length - 1
                ? tTestName.Split('_')[index]
                : "";
        }

        private static void _SetSelSramParameters(CharPlanItem planItem, VbtFunction function, bool genPmode = true, bool isCSharp = false)
        {
            //var userdefLast = planItem.TestInstanceName.Split('_').Where(x => !string.IsNullOrEmpty(x)).LastOrDefault();
            var userdefLast = planItem.TestInstanceName.Split('_').Where(x => !string.IsNullOrEmpty(x)).ToList();

            if (userdefLast.Count > 10)
            {
                if (Regex.IsMatch(userdefLast[10], "SELSRM|SELSRAM", RegexOptions.IgnoreCase))
                {
                    if (!isCSharp)
                    {
                        function.SetParamValue("SELSRAM_DSSC", "\'" + userdefLast[10]);
                    }
                }
            }

            function.SetParamValue("Vbump", "True");
            if (genPmode)
            {
                function.SetParamValue("PMode",
                    _GetCategoryFromTestName(planItem.IsOverWriteVoltage, planItem.TestInstanceName, planItem.DcCategory,
                        true));
            }

            //function.SetParamValue("Vbump",
            //    Regex.IsMatch(_GetInfoFromTestName(planItem.TestInstanceName, 10), "SelSr[a]*m", RegexOptions.IgnoreCase)
            //        ? "True"
            //        : "False");            
        }


        public static string _ConvertCategoryInAcSpec(string tTestName, string manualsymbol)
        {
            string domain = UpdateGpuDomain(tTestName.Substring(0, 3));  // SOC
            string categoryTname = tTestName.Substring(3, tTestName.Split('_').First().Length - 3);            // SOCSCCHAIN
            string freqvalue = tTestName.Split('_').Last();
            string result;

            // Mbist
            if (Regex.IsMatch(categoryTname, "(BIST|BST|BIR)", RegexOptions.IgnoreCase))
            {
                result = string.IsNullOrEmpty(freqvalue) ? domain + "Mbist" : domain + "Mbist_" + freqvalue;
            }

            // Scan
            else if (Regex.IsMatch(categoryTname, "(SA|TD)", RegexOptions.IgnoreCase))
            {
                result = string.IsNullOrEmpty(freqvalue) ? domain + "Scan" : domain + "Scan_" + freqvalue;
            }
            else
            {
                result = tTestName;
            }

            return $"{result}_{manualsymbol}";
        }

        //private static string UpdateGpuDomain(string domain)
        //{
        //    return domain.ToUpper().Equals("Gpu") ? "Gfx" : domain;
        //}

        private static string _GetCategoryFromTestName(bool isOverWriteVoltage, string tTestName, string planItemCategory, bool isUsedInPmode)
        {
            // tTestName example: DFTLH_SOC_VDDICSMCUIONA_F0100_SOCSACHAIN_ccn0_pl00_X_X_X_DSelSramSS_CZ_NV
            string result = planItemCategory;
            string[] userDefines = tTestName.Split('_');
            if (userDefines.Length < 8)
            {
                return planItemCategory;
            }

            string domain = UpdateGpuDomain(userDefines[1]);  // SOC
            string perfMode = userDefines[3];                 // F0100
            string categoryTname = userDefines[4];            // SOCSCCHAIN
            string userDef8 = userDefines[7];                 // X

            #region Retention => use tool logic search result            
            if (Regex.IsMatch(categoryTname, "RET", RegexOptions.IgnoreCase))
            {
                if (!isOverWriteVoltage)
                {
                    return isUsedInPmode
                        ? _GetRetDcCategory(domain, categoryTname, userDef8)
                        : "Mbist_" + domain + "_init_X";
                }
                else
                {
                    return isUsedInPmode ? result + ":NV" : result;
                }
            }
            #endregion

            #region Mbist
            if (Regex.IsMatch(categoryTname, "(BIST|BST|BIR)", RegexOptions.IgnoreCase))
            {
                result = _GetMbistDcCategory(domain, perfMode, categoryTname, isUsedInPmode);
            }
            #endregion

            #region Scan
            else if (Regex.IsMatch(categoryTname, "(SA|TD)", RegexOptions.IgnoreCase))
            {
                result = _GetScanDcCategory(domain, perfMode, categoryTname);
            }
            #endregion

            if (isOverWriteVoltage)
            {
                result = planItemCategory;
            }

            #region check exist in dc spec
            if (!LocalSpecs.TestProgram.DcCategoryList.Exists(p =>
                p.Equals(result.Split(':')[0], StringComparison.OrdinalIgnoreCase)))
            {
                result = planItemCategory;
            }
            #endregion

            #region judge pmode usage
            if (isUsedInPmode && !result.Contains(":"))
            {
                return result + ":NV";
            }
            else
            {
                return result;
            }
            #endregion
        }

        private static string _GetScanDcCategory(string domain, string perfMode, string categoryTname)
        {
            // (Sa|Td)(Chain)?_domain_X_(X|pmode)
            string cate = Regex.Match(categoryTname, "(?<cate>(SA|TD)(Chain)*)", RegexOptions.IgnoreCase).Groups["cate"].Value;
            var nameList = new List<string>
            {
                cate,
                domain,
                "X",
                cate.ToLower().Contains("td")
                    ? perfMode
                    : "X"
            };
            if (cate.ToLower().Contains("td"))
            {
                return "Bincut_X_X_X";
            }
            // Search again with chain removed
            if (!LocalSpecs.TestProgram.DcCategoryList.Exists(p =>
                p.Equals(string.Join("_", nameList), StringComparison.OrdinalIgnoreCase)))
            {
                nameList[0] = Regex.Replace(nameList[0], "chain", "", RegexOptions.IgnoreCase);
            }

            return string.Join("_", nameList);
        }

        private static string _GetMbistDcCategory(string domain, string pMode, string categoryTname, bool isUsedInPmode)
        {
            // MBist_[domain]_[Bist/Bira]_[pmode]:NV
            bool hasPMode = !Regex.IsMatch(pMode, @"999|F\d+|000|00100", RegexOptions.IgnoreCase);
            string voltage = hasPMode ? ":NV" : ":LV";
            pMode = hasPMode ? pMode : "VMargin3";
            if (isUsedInPmode)
            {
                if (hasPMode)
                {
                    return $"{"Bincut_X_X_X"}{voltage}";
                }

                return "Mbist_" + domain + (IsBiraItem(categoryTname) ? "_BIRA_" : "_BIST_") + pMode + voltage;
            }
            if (hasPMode)
            {
                return "Bincut_X_X_X";
            }

            return "Mbist_" + domain + (IsBiraItem(categoryTname) ? "_BIRA_" : "_BIST_") + pMode;
        }

        private static string _GetRetDcCategory(string domain, string categoryTname, string userDef8)
        {
            var retentionDictionary = new Dictionary<string, string>
            {
                {"NAP", "NRT"},
                {"NRT", "NRT"},
                {"SRT", "SRT"},
                {"SLP", "SRT"}
            };

            var nameList = new List<string>
            {
                "Mbist",
                domain,
                retentionDictionary.ContainsKey(categoryTname.ToUpper())
                    ? retentionDictionary[categoryTname.ToUpper()]
                    : !IsBiraItem(userDef8)
                        ? "ERTBIST"
                        : "ERTBIRA",
                "X"
            };

            if (!LocalSpecs.TestProgram.DcCategoryList.Exists(p =>
                p.Equals(string.Join("_", nameList), StringComparison.OrdinalIgnoreCase)))
            {
                nameList[2] = Regex.Replace(nameList[2], "BIST", "BIRA", RegexOptions.IgnoreCase);
            }

            string dcCategory = string.Join("_", nameList);
            return dcCategory + _GetNonZeroDcSelector(dcCategory);
        }

        private static string _GetNonZeroDcSelector(string dcCategory)
        {
            if (_dcCategoryHlnvDict.TryGetValue(dcCategory, out string selector))
            {
                return selector;
            }

            _dcCategoryHlnvDict[dcCategory] = ":HV";
            CategoryInSpec dcspecData = LocalSpecs.TestProgram.DcSpecDatas[0].CategoryList.FirstOrDefault
                (p => p.Name.Equals(dcCategory, StringComparison.OrdinalIgnoreCase));

            if (dcspecData == null)
            {
                return ":HV";
            }

            var valuelist = new Dictionary<string, string>
            {
                {":HV", dcspecData.Max.Trim('=')},
                {":NV", dcspecData.Typ.Trim('=')},
                {":LV", dcspecData.Min.Trim('=')}
            };

            foreach (KeyValuePair<string, string> item in valuelist
                .Where(item => double.TryParse(item.Value, out double val) && Math.Abs(val) > 1E-7))
            {
                _dcCategoryHlnvDict[dcCategory] = item.Key;
                break;
            }
            return _dcCategoryHlnvDict[dcCategory];
        }

        private static string UpdateGpuDomain(string domain)
        {
            var gpuGroup = new List<string> { "gfx", "gpu" }; // gpu key list

            if (!gpuGroup.Contains(domain.ToLower()))
            {
                return domain;
            }

            foreach (string newDomain in LocalSpecs.TestProgram.DcCategoryList
                .Where(dcCat => dcCat.Split('_').Length > 1)
                .Select(dcCat => dcCat.Split('_')[1])
                .Where(dcCat => gpuGroup.Contains(dcCat.ToLower())))
            {
                domain = newDomain;
                break;
            }
            return domain;
        }

        private static bool IsBiraItem(string flag)
        {
            return flag.ToLower().Contains("bir");
        }

        public static string ConvertCharCondition(string charCondition)
        {
            string conditionStr = charCondition.Trim(',').Replace(",", ";");
            var newList = new List<string>();
            foreach (string condition in conditionStr.Split(';'))
            {
                string[] array = condition.Split(':');

                if (array.Length < 2)
                {
                    continue;
                }

                string pinName = array[0].ToUpper();
                string forceType = Regex.Replace(array[1], "vcm", "Vicm", RegexOptions.IgnoreCase);
                if (LocalSpecs.ProgInfo.PinDic.TryGetValue(pinName, out string value))
                {
                    pinName = value;
                }

                switch (array.Length)
                {
                    case 3:
                        string forceValue = array[2];
                        newList.Add(pinName + ":" + forceType + ":" + forceValue);
                        break;

                    case 2:
                        newList.Add(pinName + ":" + forceType);
                        break;
                }
            }
            return string.Join(";", newList);
        }

        private static string _ConvertSelsrmPattern(string pattern)
        {
            //if (pattern.Contains(":")) return pattern;
            //if (Regex.IsMatch(pattern, "srmdssc", RegexOptions.IgnoreCase))
            //    return string.Format("{0}:SELSRM", pattern);
            return pattern;
        }

        public static string GetCharTmpsFlow(string tempKey, string flowTmpsName)
        {
            var resultList = new List<string> { "Flow_" + flowTmpsName };
            if (string.IsNullOrEmpty(tempKey))
            {
                return "Flow_" + flowTmpsName;
            }

            GetTMPS_temperature(tempKey, out string low, out string high);
            if (!string.IsNullOrEmpty(low))
            {
                resultList.Add("L" + low);
            }

            if (!string.IsNullOrEmpty(high))
            {
                resultList.Add("H" + high);
            }

            return string.Join("_", resultList);
        }

        public static string GetCharAdaptiveCoolingTmpsFlow(string tempKey, string flowTmpsName, string jobName)
        {

            var resultList = new List<string> { "Flow_" + flowTmpsName };
            if (string.IsNullOrEmpty(tempKey))
            {
                return "Flow_" + flowTmpsName;
            }

            GetAdaptiveCooling_temperature(jobName, out string lowTemp, out string highTemp);

            if (!string.IsNullOrEmpty(lowTemp))
            {
                resultList.Add("L" + lowTemp);
            }

            if (!string.IsNullOrEmpty(highTemp))
            {
                resultList.Add("H" + highTemp);
            }

            return string.Join("_", resultList);
        }
        public static void GetAdaptiveCooling_temperature(string job, out string low, out string high)
        {
            low = "";
            high = "";
            LocalSpecs.AdaptiveCooling.TryGetValue(job, out AdaptiveCoolingData adaptiveCoolingData);

            if (adaptiveCoolingData == null)
            {
                return;
            }

            double.TryParse(adaptiveCoolingData.TemperatureC, out double tempc);
            double.TryParse(adaptiveCoolingData.MinDeltaC, out double minDelta);
            double.TryParse(adaptiveCoolingData.MaxDeltaC, out double maxDelta);

            low = (tempc + minDelta).ToString();
            high = (tempc + maxDelta).ToString();
        }

        public static void GetAdaptiveCoolingCount(string job, out string cnt)
        {
            cnt = "";
            LocalSpecs.AdaptiveCooling.TryGetValue(job, out AdaptiveCoolingData adaptiveCoolingData);
            if (adaptiveCoolingData == null)
            {
                return;
            }

            double.TryParse(adaptiveCoolingData.TimeoutSec, out double timeOutSec);

            double timeOutCnt = timeOutSec / 0.1;
            cnt = timeOutCnt.ToString();

        }

        public static void GetTMPS_temperature(string tempKey, out string low, out string high)
        {
            low = "";
            high = "";
            string regHigh = @"(?<high>\d+)c*";
            string regLow = @"(?<low>\d+)c*";
            string regTempKeyLh = string.Join(@"\s*", new List<string> { regHigh, @"\>", "temp", @"\>", regLow });
            string regTempKeyHl = string.Join(@"\s*", new List<string> { regLow, @"\<", "temp", @"\<", regHigh });//@"(?<low>\d+c*)\s*\<\s*temp\s*\<(?<high>\d+c*)";
            string regTempKeyL1 = string.Join(@"\s*", new List<string> { regLow, @"\<", "temp" }); //@"(?<low>\d+c*)\s*\<\s*temp";
            string regTempKeyL2 = string.Join(@"\s*", new List<string> { "temp", @"\>", regLow }); //@"temp\s*\>\s*(?<low>\d+c*)";
            string regTempKeyH1 = string.Join(@"\s*", new List<string> { regHigh, @"\>", "temp" });// @"(?<high>\d+c*)\>\s*temp";
            string regTempKeyH2 = string.Join(@"\s*", new List<string> { "temp", @"\<", regHigh }); //@"temp\s*\<(?<high>\d+c*)";

            if (Regex.IsMatch(tempKey, regTempKeyLh, RegexOptions.IgnoreCase))
            {
                low = Regex.Match(tempKey, regTempKeyLh, RegexOptions.IgnoreCase).Groups["low"].Value;
                high = Regex.Match(tempKey, regTempKeyLh, RegexOptions.IgnoreCase).Groups["high"].Value;
            }
            else if (Regex.IsMatch(tempKey, regTempKeyHl, RegexOptions.IgnoreCase))
            {
                low = Regex.Match(tempKey, regTempKeyHl, RegexOptions.IgnoreCase).Groups["low"].Value;
                high = Regex.Match(tempKey, regTempKeyHl, RegexOptions.IgnoreCase).Groups["high"].Value;
            }
            else if (Regex.IsMatch(tempKey, regTempKeyL1, RegexOptions.IgnoreCase))
            {
                low = Regex.Match(tempKey, regTempKeyL1, RegexOptions.IgnoreCase).Groups["low"].Value;
            }
            else if (Regex.IsMatch(tempKey, regTempKeyL2, RegexOptions.IgnoreCase))
            {
                low = Regex.Match(tempKey, regTempKeyL2, RegexOptions.IgnoreCase).Groups["low"].Value;
            }
            else if (Regex.IsMatch(tempKey, regTempKeyH1, RegexOptions.IgnoreCase))
            {
                high = Regex.Match(tempKey, regTempKeyH1, RegexOptions.IgnoreCase).Groups["high"].Value;
            }
            else if (Regex.IsMatch(tempKey, regTempKeyH2, RegexOptions.IgnoreCase))
            {
                high = Regex.Match(tempKey, regTempKeyH2, RegexOptions.IgnoreCase).Groups["high"].Value;
            }
        }
    }
}
