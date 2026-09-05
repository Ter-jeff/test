using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.Controller;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;
using Cautogen.common.IgxlDataExtension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class InstanceGenerator
    {
        /* Property */
        private readonly bool _genCharNotUse;
        private readonly bool _genPatNotUse;
        private readonly string _enableWords;
        private readonly bool _genPmode;
        private readonly string _job;
        private readonly bool _useNewTChar;
        private readonly bool _genCSharp;
        private readonly Dictionary<string, int> _instanceDic;
        /* Constructor */
        public InstanceGenerator(InputParam inputParam)
        {
            _genCharNotUse = inputParam.GenCharNotUse;
            _genPatNotUse = inputParam.GenPatNotUse;
            _enableWords = inputParam.EnableWords;
            _genPmode = inputParam.GenPmode;
            _job = inputParam.JobName;
            _useNewTChar = inputParam.UseNewTChar;
            _genCSharp = inputParam.GenCSharp;
            _instanceDic = new Dictionary<string, int>();
        }

        /* Member function */
        public void Generate(List<CharPlanSheet> charPlanSheets)
        {
            GeneralFunc.WriteMessage("Generating instance sheets... ");

            List<InstanceRow> activeInstanceRows = _GetActiveInstanceRows(_job);

            string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.CzFolder);

            //write common instances to TestInst_Common_char
            _WriteTestInstCommon();
            _WriteTestInstCommonChar(outputFolder, charPlanSheets);

            //loop through each char plan sheet
            foreach (CharPlanSheet planSheet in charPlanSheets)
            {
                //create blank instance sheet for each char plan sheet
                var czInstSheet = new InstanceSheet("TestInst_CZ_" + planSheet.SheetName);

                if (!_genCSharp)
                {
                    WriteHeaderFooters(planSheet.SheetName, czInstSheet);
                }

                //handle hard ip sheet
                if (planSheet.IsHardIp)
                {
                    IEnumerable<CharPlanItem> missItemList = _HipInstanceHandler(planSheet, czInstSheet);
                    _MissingItemHandler(czInstSheet, missItemList);
                }
                // handle scan/mbist char sheet
                else
                {
                    InstanceSheet preScanInstSheet = _AtpgInstanceHandler(planSheet, czInstSheet, _genCSharp);
                    //export preScan inst sheet
                    //var scanFileName = Path.Combine(outputFolder, preScanInstSheet.Name + ".txt");
                    //preScanInstSheet.Write(scanFileName);
                }

                for (int i = czInstSheet.Rows.Count - 1; i >= 0; i--)
                {
                    InstanceRow czInstanceRow = czInstSheet.Rows.ElementAt(i);
                    InstanceRow activeInstanceRow = activeInstanceRows.Find(x => x.TestName.Equals(czInstanceRow.TestName, StringComparison.CurrentCultureIgnoreCase)
                                                                        && !string.IsNullOrEmpty(czInstanceRow.TestName)
                                                                        && !string.Equals(czInstSheet.Name, x.SheetName, StringComparison.OrdinalIgnoreCase));
                    if (activeInstanceRow != null)
                    {
                        if (!activeInstanceRow.GetDifferences(czInstanceRow).Any())
                        {
                            czInstSheet.Rows.Remove(czInstanceRow);
                        }
                        else
                        {
                        }
                    }
                }

                // export cz inst sheet
                string czFileName = Path.Combine(outputFolder, czInstSheet.Name + ".txt");
                czInstSheet.Write(czFileName);
                LocalSpecs.GenSheets.Add(czInstSheet);

                if (File.Exists(czFileName))
                {
                    LocalSpecs.TestProgram.JoblistSheet.AddInstanceSheet(czInstSheet.Name);
                }
            }
        }

        private List<InstanceRow> _GetActiveInstanceRows(string job)
        {
            var activeInstanceRows = new List<InstanceRow>();
            JobListSheet jobSheet = LocalSpecs.TestProgram.JoblistSheet;
            JobRow jobRow = jobSheet.Rows.Find(x => x.JobName.Equals(job, StringComparison.CurrentCultureIgnoreCase));
            if (jobRow != null)
            {
                List<string> testInstances = jobRow.TestInstances.Split(',').ToList();

                foreach (InstanceSheet instanceSheet in LocalSpecs.TestProgram.InstanceSheets)
                {
                    if (testInstances.Contains(instanceSheet.Name, StringComparer.CurrentCultureIgnoreCase))
                    {
                        activeInstanceRows.AddRange(instanceSheet.Rows);
                    }
                }
            }
            return activeInstanceRows;
        }

        private InstanceSheet _AtpgInstanceHandler(CharPlanSheet planSheet, InstanceSheet czInstSheet, bool genCSharp = false)
        {

            var preScanInstSheet = new InstanceSheet("TestInst_PreScan_" + planSheet.SheetName);
            bool enableMtd = planSheet.CharList.Any(p => !string.IsNullOrEmpty(p.Die));
            Dictionary<string, List<CharPlanItem>> mtdGroupItems = null;
            if (enableMtd)
            {
                mtdGroupItems = planSheet.CharList.GroupBy(p => p.TestInstanceName).ToDictionary(p => p.Key, p => p.ToList());

            }

            HashSet<string> recordGenerated = new HashSet<string>();

            foreach (CharPlanItem planItem in planSheet.CharList
                .Where(planItem => (_genCharNotUse || planItem.Use) &&
                                   (_genPatNotUse || SearchInfo.CheckPatUsed(planItem))))
            {
                try
                {
                    const string vbtName = VbtFunctionLib.FunctionalCharName;

                    string testInstanceName = planItem.CharShmooSetup == null
                        ? planItem.TestInstanceName + planItem.Voltage
                        : planItem.TestInstanceName + "CZ_" + planItem.Voltage;

                    string testInstanceNameHv = planItem.TestInstanceName + "Pre_HV";
                    string testInstanceNameLv = planItem.TestInstanceName + "Pre_LV";
                    if (recordGenerated.Contains(planItem.TestInstanceName))
                    {
                        continue;
                    }

                    InstanceRow czInstRow;
                    if (enableMtd && mtdGroupItems != null)
                    {
                        bool groupMtdInstance = mtdGroupItems.TryGetValue(planItem.TestInstanceName, out List<CharPlanItem> charPlanItems);
                        czInstRow = DataConvertor.ConvertPlanItemToInstanceRow(charPlanItems, vbtName, _genPmode, _useNewTChar, isCSharp: genCSharp);

                    }
                    else
                    {
                        czInstRow = DataConvertor.ConvertPlanItemToInstanceRow(planItem, vbtName, _genPmode, _useNewTChar, isCSharp: genCSharp);
                    }

                    czInstRow.TestName = testInstanceName;
                    czInstSheet.AddRow(czInstRow);
                    if (planItem.CharShmooSetup != null)
                    {
                        planItem.CharShmooSetup.Timeset = czInstRow.TimeSets;
                    }

                    planItem.InstanceName = czInstRow.TestName;
                    recordGenerated.Add(planItem.TestInstanceName);
                }
                catch (Exception)
                {
                }
            }
            return preScanInstSheet;
        }

        private IEnumerable<CharPlanItem> _HipInstanceHandler(CharPlanSheet planSheet, InstanceSheet czInstSheet)
        {
            var flowSheets = new List<string>();
            var instNotFoundPatternList = new List<string>();
            var missItemList = new List<CharPlanItem>();
            try
            {
                planSheet.CharList = _ExtendCharListForBurst(planSheet.CharList);
                foreach (CharPlanItem charRow in planSheet.CharList.Where(planItem =>
                            (_genCharNotUse || planItem.Use) && (_genPatNotUse || SearchInfo.CheckPatUsed(planItem))))
                {
                    ProcessHipInstanceRow(charRow, czInstSheet, flowSheets, instNotFoundPatternList, missItemList);
                }
            }
            catch (Exception e)
            {
                GeneralFunc.WriteMessage("Generate Instance failed " + e.Message);
            }
            if (!LocalSpecs.ProgFlowDic.ContainsKey(planSheet.SheetName))
            {
                LocalSpecs.ProgFlowDic.Add(planSheet.SheetName, flowSheets);
            }

            return missItemList;
        }

        private void ProcessHipInstanceRow(CharPlanItem charRow, InstanceSheet czInstSheet, List<string> flowSheets, List<string> instNotFoundPatternList, List<CharPlanItem> missItemList)
        {
            List<InstanceRow> instSamePayload1List = ResolveInstSamePayload1List(charRow);

            if (instSamePayload1List == null || instSamePayload1List.Count == 0)
            {
                if (!instNotFoundPatternList.Contains(charRow.Payload1))
                {
                    instNotFoundPatternList.Add(charRow.Payload1);
                    GeneralFunc.WriteMessage("Can not find production instance for pattern : " + charRow.Payload1);
                }
                missItemList.Add(charRow);
                return;
            }

            InstanceRow progInstance = _DecideWhichInstAmongSamePayload1List(instSamePayload1List, charRow);
            if (progInstance == null)
            {
                return;
            }

            charRow.InProgInstance = true;

            if (!instNotFoundPatternList.Contains(charRow.Payload1) ||
                progInstance.VbtName.ToLower() == VbtFunctionLib.DdrLpBkFunc2)
            {
                instNotFoundPatternList.Add(charRow.Payload1);
                GeneralFunc.WriteMessage("Payload \"" + charRow.Payload1 + "\" found " +
                                         instSamePayload1List.Count +
                                         " production instances, and used " + progInstance.TestName);
            }

            InstanceRow czInstRow = DataConvertor.ConvertProgInstanceToInstanceRow(charRow, progInstance, _genCSharp);
            //cz testname generate

            if (charRow.CharShmooSetup != null)
            {
                charRow.CharShmooSetup.Timeset = czInstRow.TimeSets;
            }

            if (charRow.MeasType == "HF" || charRow.MeasType == "HIO")
            {
                AssignHfHioTestName(czInstRow, charRow, progInstance);
            }
            else
            {
                czInstRow = ResolveNonHfHioTestName(czInstRow, charRow, progInstance);
                if (czInstRow == null)
                {
                    return;
                }
            }

            czInstSheet.AddRow(czInstRow);

            if (charRow.MeasType == "HIO")
            {
                AppendHioCharRow(czInstSheet, czInstRow, charRow);
            }

            if (!_instanceDic.ContainsKey(czInstRow.TestName))
            {
                _instanceDic.Add(czInstRow.TestName, 0);
            }
            else
            {
                _instanceDic[czInstRow.TestName]++;
                czInstRow.TestName = czInstRow.TestName + "_" + _instanceDic[czInstRow.TestName];
            }

            // add instance name information for generation flow table use.
            charRow.InstanceName = czInstRow.TestName;

            // add instance sheet information for generation flow table use.
            if (LocalSpecs.AllFlowStepsDic.ContainsKey(progInstance.TestName.ToLower()))
            {
                string flowSheetName = LocalSpecs.AllFlowStepsDic[progInstance.TestName.ToLower()].SheetName;
                if (!flowSheets.Contains(flowSheetName))
                {
                    flowSheets.Add(flowSheetName);
                }
            }

            UpdatePatternDictForInstance(charRow, progInstance);
        }

        private static List<InstanceRow> ResolveInstSamePayload1List(CharPlanItem charRow)
        {
            if (charRow.IsUseRtosCmd)
            {
                return ProdProg.SamePatternInstanceDict.SelectMany(p => p.Value)
                    .ToList()
                    .FindAll(p => SearchInfo.IsContains(p.TestName, charRow.Payload1));
            }

            string findKey =
                ProdProg.SamePatternInstanceDict.Where(x => x.Key.EndsWith(charRow.Payload1.ToUpper()))
                    .Select(x => x.Key)
                    .ToList()
                    .FirstOrDefault();

            if (charRow.Burst.Equals("Yes", StringComparison.CurrentCultureIgnoreCase))
            {
                return ResolveBurstInstList(charRow);
            }

            if (string.IsNullOrEmpty(findKey))
            {
                return null;
            }

            return ProdProg.SamePatternInstanceDict.TryGetValue(findKey, out List<InstanceRow> value)
                ? value.Where(x => !string.IsNullOrEmpty(x.TestName)).ToList()
                : null;
        }

        private static List<InstanceRow> ResolveBurstInstList(CharPlanItem charRow)
        {
            List<InstanceRow> instSamePayload1List = null;
            var multipleItem =
                ProdProg.SamePatternInstanceDict.Where(x => x.Key.Contains(charRow.MappingPatternSet.ToUpper()))
                    .Select(x => x.Key)
                    .ToList();

            if (!multipleItem.Any())
            {
                return null;
            }

            foreach (string item in multipleItem.Where(item => ProdProg.SamePatternInstanceDict.ContainsKey(item)))
            {
                //instSamePayload1List = ProdProg.SamePatternInstanceDict[item].Where(
                //    x => !String.IsNullOrEmpty(x.TestName) && Regex.IsMatch(x.TestName, "MULTIPLE"))
                //    .ToList();
                instSamePayload1List = ProdProg.SamePatternInstanceDict[item].Where(
                    x => !string.IsNullOrEmpty(x.TestName))
                    .ToList();
                if (instSamePayload1List.Any())
                {
                    break;
                }
            }
            return instSamePayload1List;
        }

        private static void AssignHfHioTestName(InstanceRow czInstRow, CharPlanItem charRow, InstanceRow progInstance)
        {
            if (charRow.Burst.ToLower() != "yes" && progInstance.TestName.ToUpper().Contains("INIT"))
            {
                string progtestName =
                    Regex.Match(progInstance.TestName, @"(?<sheet>(CZ|PP)_\w+)$").Groups["sheet"].ToString();
                string charTestName =
                    Regex.Match(charRow.TestInstanceName, @"(?<sheet>\w+)" + charRow.BlockName).Groups["sheet"].ToString();
                czInstRow.TestName = charTestName + progtestName;
            }
            else
            {
                czInstRow.TestName = charRow.TestInstanceName + "CZ_" + charRow.Voltage;
            }
        }

        // Returns null to signal `continue` (DdrLpBkFunc2 special case where TestName is already cached).
        private InstanceRow ResolveNonHfHioTestName(InstanceRow czInstRow, CharPlanItem charRow, InstanceRow progInstance)
        {
            czInstRow = _ClearSweepInfoInTestProg(czInstRow);

            //clear sweep information if test program has alreadey exist 2017/7/12
            if (charRow.TestInstanceName.Contains("SweepCode"))
            {
                czInstRow = _ConvertToSweepCodeInstanceRow(czInstRow, charRow, progInstance);
            }
            else if (charRow.TestInstanceName.Contains("SweepVoltage"))
            {
                czInstRow = _ConvertToSweepVoltageInstanceRow(czInstRow, charRow, progInstance);
            }
            else
            {
                czInstRow.TestName = _GetTestInstName(charRow, progInstance);
            }

            if (czInstRow.VbtName.Equals("Opt_DdrLpBkFunc2", StringComparison.OrdinalIgnoreCase))
            {
                czInstRow.TestName = charRow.TestInstanceName + "CZ_" + charRow.Voltage;
                charRow.InstanceName = czInstRow.TestName;
                if (!DataConvertor.SpecialHacInstanceRows.ContainsKey(progInstance.TestName))
                {
                    DataConvertor.SpecialHacInstanceRows.Add(progInstance.TestName, czInstRow.TestName);
                }
                else
                {
                    return null;
                }
            }
            return czInstRow;
        }

        private static void AppendHioCharRow(InstanceSheet czInstSheet, InstanceRow czInstRow, CharPlanItem charRow)
        {
            InstanceRow cpCharInstRow = czInstRow.Copy();
            cpCharInstRow.TestName = charRow.TestInstanceName + "CZ_" + charRow.Voltage + "_Char";

            VbtFunction function = BasMain.VbtFunctionLib.GetFunctionByName(cpCharInstRow.VbtName);
            function.ArgList = cpCharInstRow.Args;
            string prePat = charRow.CharCondition;
            if (charRow.CharShmooSetup != null)
            {
                prePat += ";CharSetName:" + charRow.CharShmooSetup.ShmooSetupName;
            }

            function.SetParamValue("interpose_prepat", prePat.Trim(';'));
            cpCharInstRow.Args = function.ArgList;
            czInstSheet.AddRow(cpCharInstRow);
        }

        // add pattern information for flow step
        private static void UpdatePatternDictForInstance(CharPlanItem charRow, InstanceRow progInstance)
        {
            string subInstanceName = progInstance.TestName.Replace("_HV", "").Replace("_NV", "").Replace("_LV", "").ToLower();

            if (LocalSpecs.PatternDic.ContainsKey(subInstanceName))
            {
                return;
            }

            if (charRow.UsedPatterns.Count(x => !string.IsNullOrEmpty(x)) > 1 &&
                !charRow.Burst.ToLower().Equals("yes"))
            {
                if (charRow.UsedPatterns.Exists(
                        x => !string.IsNullOrEmpty(x) && SearchInfo.IsContains(subInstanceName, x)))
                {
                    LocalSpecs.PatternDic.Add(subInstanceName, charRow.UsedPatterns.FirstOrDefault(
                            x => !string.IsNullOrEmpty(x) && SearchInfo.IsContains(subInstanceName, x))
                        .ToLower());
                }
            }
            else
            {
                LocalSpecs.PatternDic.Add(subInstanceName, charRow.Payload1.ToLower());
            }
        }

        private void _MissingItemHandler(InstanceSheet czInstSheet, IEnumerable<CharPlanItem> missItemList)
        {
            if (missItemList.Count() != 0)
            {
                // add a blank row before missing items
                czInstSheet.AddRow(new InstanceRow());
            }

            foreach (CharPlanItem planItem in missItemList)
            {
                HardIpReference patInfo = SearchInfo.GetHardIpInfo(planItem.Payload1);
                //var testInstanceName = planItem.BlockName + "_" + planItem.Payload1 + "_CZ_" + planItem.Voltage;
                var nameList = new List<string>();
                nameList.Add(planItem.BlockName);
                nameList.Add(planItem.Payload1);
                nameList.Add(planItem.Type.Replace("-", "_"));
                nameList.Add("");
                string testInstanceName = string.Join("_", nameList) + "CZ_" + planItem.Voltage;
                if (patInfo == null || patInfo.MeasSeqStr == "")
                {
                    string vbtName = BasMain.FuntionalUpdated.FunctionName;
                    //if (planItem.MeasType == "HF" || planItem.MeasType == "HIO")
                    //{
                    //    testInstanceName = planItem.BlockName + "_" + planItem.TestInstanceName + "CZ_" + planItem.Voltage;
                    //}
                    InstanceRow functionalInstanceRow = DataConvertor.ConvertPlanItemToInstanceRow(planItem, vbtName, _genPmode);
                    functionalInstanceRow.TestName = testInstanceName;
                    functionalInstanceRow.ColumnA += ",MissingPayloadInProduction";
                    czInstSheet.AddRow(functionalInstanceRow);
                }
                else
                {
                    string vbtName;
                    if (patInfo.MeasFStr != "" || patInfo.MeasVPowerPinList.Count > 0 ||
                        patInfo.MeasIPowerPinList.Count > 0)
                    {
                        vbtName = VbtFunctionLib.VifName;
                    }
                    else
                    {
                        vbtName = VbtFunctionLib.VirName;
                    }

                    InstanceRow newInstanceRow = DataConvertor.ConvertPlanItemToInstanceRow(planItem, vbtName, _genPmode);
                    newInstanceRow.TestName = testInstanceName;
                    newInstanceRow.ColumnA += ",MissingPayloadInProduction";
                    czInstSheet.AddRow(newInstanceRow);
                }
                //Added instance name information for generation flow table use.
                planItem.InstanceName = testInstanceName;
            }
        }

        private InstanceRow _DecideWhichInstAmongSamePayload1List(IEnumerable<InstanceRow> instanceList, CharPlanItem planItem)
        {
            return instanceList.OrderByDescending(a =>
            {
                int order = 0;
                // special handle DDR vbt
                if (a.VbtName.Equals(VbtFunctionLib.DdrLpBkFunc2, StringComparison.OrdinalIgnoreCase)
                    && a.Args[1].Equals(planItem.Payload2, StringComparison.OrdinalIgnoreCase))
                {
                    order += 100;
                }
                // check if match sub-block
                if (_CheckSubBlock(planItem, a))
                {
                    order += 10;
                }

                if (_CheckBlockName(planItem, a))
                {
                    order += 10;
                }

                // prefer NV > HV > LV
                string testName = a.TestName.ToUpper();

                if (SearchInfo.IsContains(testName, planItem.SheetName))
                {
                    order += 5;
                }

                if (testName.EndsWith("NV"))
                {
                    order += 3;
                }

                if (testName.EndsWith("HV"))
                {
                    order += 2;
                }

                if (testName.EndsWith("LV"))
                {
                    order += 1;
                }

                return order;

            }).FirstOrDefault();
        }

        private void WriteHeaderFooters(string sheetName, InstanceSheet czInstSheet)
        {
            czInstSheet.AddRow(_CreateHeader(sheetName.Replace(" ", "_") + "_Header", sheetName));
            czInstSheet.AddRow(_CreateFooter(sheetName.Replace(" ", "_") + "_Footer", sheetName));
            czInstSheet.AddRow(_CreateHeader(sheetName.Replace(" ", "_") + "_HVCC_Header", sheetName + "_HVCC"));
            czInstSheet.AddRow(_CreateHeader(sheetName.Replace(" ", "_") + "_LVCC_Header", sheetName + "_LVCC"));
            czInstSheet.AddRow(_CreateFooter(sheetName.Replace(" ", "_") + "_HVCC_Footer", sheetName + "_HVCC"));
            czInstSheet.AddRow(_CreateFooter(sheetName.Replace(" ", "_") + "_LVCC_Footer", sheetName + "_LVCC"));
        }

        // create a header instance
        private InstanceRow _CreateHeader(string instName, string printStr)
        {
            return new InstanceRow
            {
                TestName = instName,
                VbtName = "Print_Header",
                VbtType = "VBT",
                ArgList = "PrintInfo",
                Args = new List<string> { printStr },
            };
        }

        // create a footer instance
        private InstanceRow _CreateFooter(string instName, string printStr)
        {
            return new InstanceRow
            {
                TestName = instName,
                VbtName = "Print_Footer",
                VbtType = "VBT",
                ArgList = "PrintInfo",
                Args = new List<string> { printStr },
            };
        }

        private void _CheckThenAdd(InstanceSheet instSheet, InstanceRow instRow)
        {
            if (!ProdProg.AllTestInstances.Exists(x => x.TestName.Equals(instRow.TestName, StringComparison.CurrentCultureIgnoreCase) &&
                                                      !x.SheetName.Equals(instSheet.Name, StringComparison.CurrentCultureIgnoreCase)))
            {
                instSheet.AddRow(instRow);
            }
        }

        private void _WriteTestInstCommon()
        {
            InstanceSheet commonInstSheet = LocalSpecs.TestProgram.InstanceSheets.FirstOrDefault(p => p.Name
            .Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            var onProgramStartedBinOutFunctionInstance = new InstanceRow() { TestName = "onProgramStartedBinOutFunction", VbtType = "VBT", VbtName = "onProgramStartedBinOutFunction" };
            if (commonInstSheet != null)
            {
                if (BasMain.VbtFunctionLib.IsFunctionExist(onProgramStartedBinOutFunctionInstance.VbtName))
                {
                    _CheckThenAdd(commonInstSheet, onProgramStartedBinOutFunctionInstance);
                }
            }
            else
            {
                return;
            }
            string outputFolder = LocalSpecs.InputParam.GenTxtOnly
                ? LocalSpecs.OutputFolder
                : Path.Combine(LocalSpecs.OutputFolder, ConstData.CommonFolder, "Common_Sheets");
            string outputPath = Path.Combine(outputFolder, commonInstSheet.Name + ".txt");
            commonInstSheet.Write(outputPath);
            LocalSpecs.GenSheets.Add(commonInstSheet);
        }

        //Create instance sheet "TestInst_Common_char"
        private void _WriteTestInstCommonChar(string outputFolder, List<CharPlanSheet> charPlanSheets)
        {
            var commonCharInstSheet = new InstanceSheet("TestInst_Common_char");
            //_CheckThenAdd(commonInstSheet, new InstanceRow { TestName = ConstData.DisableCzMode, VbtName = ConstData.TpModeOffModule, VbtType = "VBT" });

            if (_genCSharp)
            {
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = ConstData.EnableCzMode, VbtName = ConstData.TpModeOnModuleCSharp, VbtType = ".NET" });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = ConstData.DisableCzMode, VbtName = ConstData.TpModeOffModuleCSharp, VbtType = ".NET" });

                //Add test instance to common test instance sheet

                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Hole_Ratio_Within_Spec", VbtName = "IgxlWrapper.CoreTestLibrary.Char.FunctionalTestCharMain.CheckCharErrorCount", VbtType = ".NET", ArgList = "shmooAbnormalType,shmooAbnormalRatio_Hilimt", Args = new List<string> { "shmoo_hole", "0.1" } });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Allfail_Ratio_Within_Spec", VbtName = "IgxlWrapper.CoreTestLibrary.Char.FunctionalTestCharMain.CheckCharErrorCount", VbtType = ".NET", ArgList = "shmooAbnormalType,shmooAbnormalRatio_Hilimt", Args = new List<string> { "all_fail", "0.1" } });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Alarm_Ratio_Within_Spec", VbtName = "IgxlWrapper.CoreTestLibrary.Char.FunctionalTestCharMain.CheckCharErrorCount", VbtType = ".NET", ArgList = "shmooAbnormalType,shmooAbnormalRatio_Hilimt", Args = new List<string> { "alarm", "0.1" } });
                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Disable_Shmoo_Abnormal_Counter", VbtName = "DisableShmooAbnormalCounter", VbtType = ".NET" });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Char_Setup_Gating", VbtName = "IgxlWrapper.CoreTestLibrary.Char.FunctionalTestCharMain.CharSetupGating", VbtType = ".NET" });

            }
            else
            {
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = ConstData.EnableCzMode, VbtName = ConstData.TpModeOnModule, VbtType = "VBT" });

                //Add test instance to common test instance sheet
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Hole_Ratio_Within_Spec", VbtName = "CheckCharErrorCount", VbtType = "VBT", ArgList = "shmoo_abnormal_type,shmoo_abnormal_ratio_hi_lim", Args = new List<string> { "shmoo_hole", "0.1" } });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Allfail_Ratio_Within_Spec", VbtName = "CheckCharErrorCount", VbtType = "VBT", ArgList = "shmoo_abnormal_type,shmoo_abnormal_ratio_hi_lim", Args = new List<string> { "all_fail", "0.1" } });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Alarm_Ratio_Within_Spec", VbtName = "CheckCharErrorCount", VbtType = "VBT", ArgList = "shmoo_abnormal_type,shmoo_abnormal_ratio_hi_lim", Args = new List<string> { "alarm", "0.1" } });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Disable_Shmoo_Abnormal_Counter", VbtName = "DisableShmooAbnormalCounter", VbtType = "VBT" });
                _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Enable_Shmoo_Abnormal_Counter", VbtName = "EnableShmooAbnormalCounter", VbtType = "VBT" });

                if (!string.IsNullOrEmpty(_enableWords))
                {
                    _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "SetEnableWords", VbtName = "SetEnableWords", VbtType = "VBT" });
                    _CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "PrintEnableWords", VbtName = "PrintEnableWords", VbtType = "VBT" });
                    _CheckThenAdd(commonCharInstSheet, _CreateHeader("PrintEnableWords_Header", "PrintEnableWords"));
                    _CheckThenAdd(commonCharInstSheet, _CreateFooter("PrintEnableWords_Footer", "PrintEnableWords"));
                }
            }

            _AddnWireForSwitch(charPlanSheets, commonCharInstSheet);

            //Write sheet if rows > 0
            if (commonCharInstSheet.Rows.Count > 0)
            {
                //Export sheet if not already exist
                string exportPath = Path.Combine(outputFolder, commonCharInstSheet.Name + ".txt");
                if (!File.Exists(exportPath))
                {
                    commonCharInstSheet.Write(exportPath);
                    LocalSpecs.GenSheets.Add(commonCharInstSheet);
                }

                LocalSpecs.TestProgram.JoblistSheet.AddInstanceSheet(commonCharInstSheet.Name);
            }
        }

        private void _AddnWireForSwitch(List<CharPlanSheet> charPlanSheets, InstanceSheet commonInstSheet)
        {
            var frcSet = charPlanSheets.SelectMany(p => p.CharList)
                .Where(p => !string.IsNullOrEmpty(p.FreeRunningClock)).Select(p => p.FreeRunningClock).Distinct().ToList();
            InstanceSheet commonInsts = LocalSpecs.TestProgram.InstanceSheets.FirstOrDefault(p => p.Name
            .Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            if (commonInsts == null)
            {
                return;
            }
            foreach (string frc in frcSet)
            {
                if (frc.Split(':').Length < 2)
                {
                    continue;
                }
                string frcName = frc.Split(':')[0].ToLower();
                if (decimal.TryParse(frc.Split(':')[1], out decimal frcFreq))
                {
                    foreach (InstanceRow commInst in commonInsts.Rows)
                    {
                        if (commInst.TestName.ToLower().Contains("freerunclk") && commInst.TestName.ToLower().Replace("_", "").Contains(frcName))
                        {
                            InstanceRow frcNewInst = commInst.Copy();
                            if (!string.IsNullOrEmpty(frcNewInst.AcCategory))
                            {
                                frcNewInst.AcCategory = string.Format($"{frcNewInst.AcCategory}_{(frcFreq / 1000000).ToString().Replace(".", "p")}");
                            }
                            frcNewInst.TestName = string.Format($"{frcNewInst.TestName}_{(frcFreq / 1000000).ToString().Replace(".", "p")}MHz");
                            int freqArgIndex = frcNewInst.ArgList.Split(',').ToList().FindIndex(s => s.Equals("freq", StringComparison.OrdinalIgnoreCase));
                            if (freqArgIndex != -1 && frcNewInst.Args.Count > freqArgIndex)
                            {
                                frcNewInst.Args[freqArgIndex] = frcFreq.ToString();
                            }
                            _CheckThenAdd(commonInstSheet, frcNewInst);
                        }
                    }
                }
            }
        }

        private InstanceRow _ConvertToSweepCodeInstanceRow(InstanceRow czinstrow, CharPlanItem charitem, InstanceRow progInstance)
        {
            czinstrow.TestName = _GetTestInstName(charitem, progInstance);
            //progInstance.TestName.Replace("_HV", "").Replace("_NV", "").Replace("_LV", "") +
            //                    "_CZ_" + "Sweep_"+ charitem.Voltage;
            string flowForLoopIntegerName = "SrcCodeIndx;";
            List<IGLinkProcessor.DataStructure.ShmooData.ShmooPin> tempShmoopin = charitem.CharShmooSetup.ShmooPins.FindAll(s => s.SweepType == "SweepCode");
            foreach (IGLinkProcessor.DataStructure.ShmooData.ShmooPin shmooPin in tempShmoopin)
            {
                string pinname = shmooPin.SweepPinName;
                string start = shmooPin.StartPoint;
                string step = shmooPin.StepSize;
                string width = _GetSendBitLength(pinname, charitem).ToString(CultureInfo.InvariantCulture);
                flowForLoopIntegerName += pinname + ":" + width + ":" + start + ":" + step + ";";
            }
            VbtFunction function = BasMain.VbtFunctionLib.GetFunctionByName(czinstrow.VbtName);
            function.ArgList = czinstrow.Args;
            function.SetParamValue("DigSrc_FlowForLoopIntegerName", flowForLoopIntegerName.Trim(';'));
            czinstrow.Args = function.ArgList;
            return czinstrow;
        }

        private int _GetSendBitLength(string sendBitName, CharPlanItem planItem)
        {
            HardIpReference patInfo = SearchInfo.GetHardIpInfo(planItem.Payload1);
            List<string> srcBitNameList = patInfo.SendBitName.Split('+').ToList();
            List<string> srcBitStrList = patInfo.SendBitStr.Split('+').ToList();

            if (sendBitName.Equals("repeat") && srcBitNameList.Count > 0)
            {
                return int.Parse(srcBitStrList[0].Split('_')[1]);
            }

            for (int i = 0; i < srcBitNameList.Count; i++)
            {
                if (srcBitNameList[i].Equals(sendBitName, StringComparison.OrdinalIgnoreCase))
                {
                    return int.Parse(srcBitStrList[i].Split('_')[1]);
                }
            }
            return 0;
        }

        private InstanceRow _ClearSweepInfoInTestProg(InstanceRow czinstrow)
        {
            VbtFunction function = BasMain.VbtFunctionLib.GetFunctionByName(czinstrow.VbtName);
            function.ArgList = czinstrow.Args;
            function.SetParamValue("DigSrc_FlowForLoopIntegerName", "");
            czinstrow.Args = function.ArgList;
            return czinstrow;
        }

        private InstanceRow _ConvertToSweepVoltageInstanceRow(InstanceRow czinstrow, CharPlanItem charitem, InstanceRow progInstance)
        {
            //czinstrow.TestName = progInstance.TestName.Replace("_HV", "").Replace("_NV", "").Replace("_LV", "") + "_CZ_" + "SweepVoltage_" + charitem.Voltage;
            czinstrow.TestName = czinstrow.TestName = _GetTestInstName(charitem, progInstance);
            return czinstrow;
        }

        private string _GetTestInstName(CharPlanItem charitem, InstanceRow progInstance)
        {
            var namingList = new List<string>();
            namingList.Add(progInstance.TestName.Split('_')[0]);
            namingList.Add(progInstance.TestName.Split('_')[1]);
            namingList.Add(charitem.Payload1);
            namingList.Add(charitem.Type.Replace("-", "_"));

            if (charitem.TestInstanceName.ToUpper().Contains("RING"))
            {
                string userdef9 = DataConvertor.GetInfoFromTestName(charitem.TestInstanceName, 10);
                namingList.Add(userdef9);
            }

            namingList.Add("NV");
            //var nameSgmts = progInstance.Name.Split('_')[0] + "_" + progInstance.Name.Split('_')[1];
            string name = charitem.IsUseRtosCmd ? progInstance.TestName : string.Join("_", namingList);
            string postfix = "";

            if (charitem.TestInstanceName.Contains("SweepCode"))
            {
                postfix = "Sweep_";
            }
            else if (charitem.TestInstanceName.Contains("SweepVoltage"))
            {
                postfix = "SweepVoltage_";
            }

            string testName = name
                .Replace("_HV", "").Replace("_NV", "").Replace("_LV", "")
                + "_CZ_" + postfix + charitem.Voltage;

            return testName;
        }

        private bool _CheckSubBlock(CharPlanItem charitem, InstanceRow czInstance)
        {
            /* Check if match charRow.Type = "1A" and czInstance.TestName = "PCIE_1A_..."*/
            if (charitem.Type == "X")
            {
                return false;
            }

            string patternName = czInstance.Args[0].ToUpper();
            return czInstance.TestName.ToUpper()
                .Split(new[] { patternName }, StringSplitOptions.RemoveEmptyEntries)[0]
                .Contains("_" + charitem.Type + "_");
        }

        private bool _CheckBlockName(CharPlanItem charitem, InstanceRow czInstance)
        {
            /* Check if match charRow.Type = "1A" and czInstance.TestName = "PCIE_1A_..."*/
            if (charitem.Type == "X")
            {
                return false;
            }

            string patternName = czInstance.Args[0].ToUpper();
            return czInstance.TestName.ToUpper()
                .Split(new[] { patternName }, StringSplitOptions.RemoveEmptyEntries)[0]
                .Contains(charitem.BlockName + "_");
        }

        private List<CharPlanItem> _ExtendCharListForBurst(List<CharPlanItem> charList)
        {
            var output = new List<CharPlanItem>();
            foreach (CharPlanItem charRow in charList)
            {
                if (charRow.UsedPatterns.Count(x => !string.IsNullOrEmpty(x)) > 1
                    && !charRow.Burst.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                {
                    string initPatSet = _IsInitPatternSet(charRow);
                    if (!string.IsNullOrEmpty(initPatSet))
                    {
                        output.Add(new CharPlanItem(charRow) { Payload1 = initPatSet, ExtendInit = true });
                        output.Add(charRow);
                    }
                    else
                    {
                        output.AddRange(charRow.UsedPatterns
                            .Where(x => !string.IsNullOrEmpty(x))
                            .Select(x => new CharPlanItem(charRow) { Payload1 = x }));
                    }
                }
                else
                {
                    output.Add(charRow);
                }
            }
            return output;
        }

        private string _IsInitPatternSet(CharPlanItem charRow)
        {
            string initPatSetName = "";
            IEnumerable<KeyValuePair<string, List<string>>> searchPatternSet = ProdProg.AllPatSetRows.Where(x => x.Value.SequenceEqual(charRow.UsedInits));

            if (searchPatternSet.Any())
            {
                initPatSetName = searchPatternSet.First().Key;
            }

            if (!string.IsNullOrEmpty(initPatSetName))
            {
                if (ProdProg.SamePatternInstanceDict.Keys.Any(x => string.Equals(x, initPatSetName, StringComparison.OrdinalIgnoreCase)))
                {
                    return initPatSetName;
                }
            }
            return "";
        }
    }
}
