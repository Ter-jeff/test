using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure.ShmooData;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;
using Cautogen.common.ReaderWriter.Reader.InputDataBase;
using Cautogen.common.ReaderWriter.Reader.InputReader;
using Cautogen.Utility;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using LogLib.Utility;

using OfficeOpenXml;

using ExcelOperation = Cautogen.AutoCZ.CharPostProcessor.Utility.ExcelOp.ExcelOperation;

namespace Cautogen.AutoCZ.CharPostProcessor.InputReader
{
    public class CharPlanReader
    {
        /* Member function */
        public static List<CharPlanSheet> Read(string fileName)
        {
            GeneralFunc.WriteMessage("Reading CharPlan " + fileName);
            LogHelper.Info("Reading CharPlan " + fileName);

            try
            {
                var charSheetList = new List<CharPlanSheet>();

                LocalSpecs.PatSetAll = new PatSetSheet("PatSetMerge");
                ExcelWorkbook wb = new ExcelPackage(new FileInfo(fileName)).Workbook;
                foreach (ExcelWorksheet sheet in wb.Worksheets)
                {
                    GeneralFunc.WriteMessage("Reading sheet " + sheet.Name + "...");
                    switch (sheet.Name)
                    {
                        case "EmaMapping":
                            LocalSpecs.EmaMappingItems = ReadEmaMappingTalbe(sheet);
                            break;

                        case "PatternDashBoard":
                            //ReadCharDashBoard(sheet);
                            break;
                        case "PatSet_AI":
                            ReadPatSet_AI(sheet);
                            break;
                        case "DFC_List":
                            if (sheet.Dimension != null)
                            {
                                LocalSpecs.DFCSheet = sheet;
                            }
                            break;
                        case "AdaptiveCooling":
                            LocalSpecs.AdaptiveCooling = new AdaptiveCoolingReader().Read(sheet);
                            break;
                        case "timesettings":
                            LocalSpecs.OptionalTimesettings = sheet;
                            break;
                        default:
                            CharPlanSheet charSheet = ReadSheet(sheet);
                            if (charSheet.CharList.Count > 0)
                            {
                                charSheetList.Add(charSheet);
                            }

                            break;
                    }

                }
                return charSheetList;
            }
            catch (Exception ex)
            {
                throw new Exception("Reading CharPlan failed! " + ex.Message);
            }
        }

        public static Dictionary<string, PatternData> ReadPatternDashBoard(string fileName)
        {
            GeneralFunc.WriteMessage("Reading DashBoard ");

            try
            {
                ExcelWorkbook wb = new ExcelPackage(new FileInfo(fileName)).Workbook;
                foreach (ExcelWorksheet sheet in wb.Worksheets)
                {
                    GeneralFunc.WriteMessage("Reading sheet " + sheet.Name + "...");
                    if (sheet.Name == "PatternDashBoard" && sheet.Dimension != null)
                    {
                        var patternList = new Dictionary<string, PatternData>();
                        for (int rowindex = 1; rowindex <= sheet.Dimension.Rows; rowindex++)
                        {
                            var item = new PatternData
                            {
                                PatternName = sheet.Cells[rowindex, 1].Text,
                                Use = sheet.Cells[rowindex, 2].Text,
                                FileVersion = sheet.Cells[rowindex, 3].Text,
                                TimesetVersion = sheet.Cells[rowindex, 4].Text
                            };
                            patternList.Add(item.PatternName, item);
                        }

                        return patternList;
                    }

                }
                return new Dictionary<string, PatternData>();
            }
            catch (Exception ex)
            {
                throw new Exception("Reading CharPlan PatternDashBoard failed! " + ex.Message);
            }
        }

        private static CharPlanSheet ReadSheet(ExcelWorksheet sheet)
        {
            var newCharSheet = new CharPlanSheet(sheet.Name);

            #region Initial HeaderIndex

            Dictionary<string, int> headerOrder = ExcelOperation.GetHeaderOrder(sheet);
            int blockNameIndex = ExcelOperation.GetHeaderIndex(headerOrder, "block name");
            int descripIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Description");
            int typeIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Type");
            int namingIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Naming Selection");
            int powerRunIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Power Run Scenario");
            int waitIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Wait");
            int useIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Use/Not Use");
            int enablewordIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Enable");

            int instanceNameIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Test Instance Name");
            int voltageIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Voltage");
            int dcSelectorIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DC Selector");
            int acSelectorIndex = ExcelOperation.GetHeaderIndex(headerOrder, "AC Selector");
            int acCategoryIndex = ExcelOperation.GetHeaderIndex(headerOrder, "AC Category");
            int dcCategoryIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DC Category");
            int levelsIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Levels");
            int timesetIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Timeset");
            int initPatternsIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Init Patterns");
            int payloadPatternsIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Payload Patterns");
            int shmooSetupIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Shmoo Setup Name");
            int xsweep1Index = ExcelOperation.GetHeaderIndex(headerOrder, "X Sweep 1");
            int ysweep1Index = ExcelOperation.GetHeaderIndex(headerOrder, "Y Sweep 1");
            int zsweep1Index = ExcelOperation.GetHeaderIndex(headerOrder, "Z Sweep 1");
            int searchMethodIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Search Method");
            int htolIndex = ExcelOperation.GetHeaderIndex(headerOrder, "HTOL");
            int conditionIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Char_Condition");
            int ttrIndex = ExcelOperation.GetHeaderIndex(headerOrder, "TTR");
            int hardIpIndex = ExcelOperation.GetHeaderIndex(headerOrder, "HardIP");
            int ipuse1Index = ExcelOperation.GetHeaderIndex(headerOrder, "IP Use_Test Items");
            int ipuse2Index = ExcelOperation.GetHeaderIndex(headerOrder, "IP Use_Description");
            int ipuse3Index = ExcelOperation.GetHeaderIndex(headerOrder, "IP Use_ProgramTestName");
            int ipuse4Index = ExcelOperation.GetHeaderIndex(headerOrder, "IP Use_USL");
            int ipuse5Index = ExcelOperation.GetHeaderIndex(headerOrder, "IP Use_LSL");
            int ipuse6Index = ExcelOperation.GetHeaderIndex(headerOrder, "IP Use_Units");
            int programTestName1Index = ExcelOperation.GetHeaderIndex(headerOrder, "ProgramTestName_1");
            int programTestName2Index = ExcelOperation.GetHeaderIndex(headerOrder, "ProgramTestName_2");
            int srcAssignmentIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DigSrcAssignment");
            int isPatternExistIndex = ExcelOperation.GetHeaderIndex(headerOrder, "PatternNotExist");
            int selSrmSendBitIndex = ExcelOperation.GetHeaderIndex(headerOrder, "SelSRMSentBit");
            int rtosCmdIndex = ExcelOperation.GetHeaderIndex(headerOrder, "RtosUseCmd");
            int suspendDatalogIndex = ExcelOperation.GetHeaderIndex(headerOrder, "SuspendDatalog");
            int harvfstpIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Harv_FSTP");
            int siteflagIndex = ExcelOperation.GetHeaderIndex(headerOrder, "SiteFlag");
            int failflagIndex = ExcelOperation.GetHeaderIndex(headerOrder, "FailFlag");
            int manualAcIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Manual_AC");
            int failInfoIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Fail_Info");
            int envIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Environment");
            int burstIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Burst");
            int isDsscReverseIndex = ExcelOperation.GetHeaderIndex(headerOrder, "IsEMA_Data_Reverse");
            int manualACfromTimesetIndex = ExcelOperation.GetHeaderIndex(headerOrder, "ManualACfromTimeset");
            int shiftFreqIndex = ExcelOperation.GetHeaderIndex(headerOrder, "ShiftFreq");
            int bypassShmooHoleIndex = ExcelOperation.GetHeaderIndex(headerOrder, "BypassShmooHole");
            int retentionRampIndex = ExcelOperation.GetHeaderIndex(headerOrder, "RetentionRamp");
            int digSrcBitSizeIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DigSrcBitSize");
            int digSrcSegIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DigSrcSeg");
            int digSrcPinIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DigSrcPin");
            int digSrcEqIndex = ExcelOperation.GetHeaderIndex(headerOrder, "DigSrcEQ");
            int oneTimeInitIndex = ExcelOperation.GetHeaderIndex(headerOrder, "OneTimeINIT");
            int mappingPatternSetIndex = ExcelOperation.GetHeaderIndex(headerOrder, "MappingPatternSet");
            int applyVoltageFromBinCutIndex = ExcelOperation.GetHeaderIndex(headerOrder, "ApplyVoltageFromBinCut");
            int freeRunningClockIndex = ExcelOperation.GetHeaderIndex(headerOrder, "FreeRunningClock");
            int userFunctionIndex = ExcelOperation.GetHeaderIndex(headerOrder, "UserFunction");
            int harvestPinGrpOtherFail = ExcelOperation.GetHeaderIndex(headerOrder, "HarvestPinGrpOtherFail");
            int enableCoreHarvest = ExcelOperation.GetHeaderIndex(headerOrder, "EnableCoreHarvest");
            int enableCoreMask = ExcelOperation.GetHeaderIndex(headerOrder, "EnableCoreMask");
            int pinGrpSpecifyMask = ExcelOperation.GetHeaderIndex(headerOrder, "PinGrpSpecifyMask");
            int ssnSpecifyMask = ExcelOperation.GetHeaderIndex(headerOrder, "SSNSpecifyMask");
            int adaptiveCooling = ExcelOperation.GetHeaderIndex(headerOrder, "AdaptiveCooling");
            int selsramUserDef9 = ExcelOperation.GetHeaderIndex(headerOrder, "SelsrmUserDef9");
            int stageCp1 = ExcelOperation.GetHeaderIndex(headerOrder, "StageCP1");
            int stageCp2 = ExcelOperation.GetHeaderIndex(headerOrder, "StageCP2");
            int stageFt1 = ExcelOperation.GetHeaderIndex(headerOrder, "StageFT1");
            int stageFt2 = ExcelOperation.GetHeaderIndex(headerOrder, "StageFT2");
            int die = ExcelOperation.GetHeaderIndex(headerOrder, "DIE");
            #endregion

            newCharSheet.IsHardIp = ExcelOperation.GetCellValue(sheet, 3, hardIpIndex).ToLower() == "y" || ExcelOperation.GetCellValue(sheet, 3, hardIpIndex).ToLower() == "yes";

            for (int row = 3; row <= sheet.Dimension.End.Row; row++)
            {
                var charItem = new CharPlanItem { RowNum = row };
                var allPatternList = new List<string>();
                charItem.BlockName = ExcelOperation.GetCellValue(sheet, row, blockNameIndex);
                //if (charItem.BlockName == "") continue;
                charItem.SheetName = sheet.Name;
                charItem.Description = ExcelOperation.GetCellValue(sheet, row, descripIndex);
                charItem.EnableWord = ExcelOperation.GetCellValue(sheet, row, enablewordIndex);
                charItem.Type = ExcelOperation.GetCellValue(sheet, row, typeIndex);
                charItem.NamingSelection = ExcelOperation.GetCellValue(sheet, row, namingIndex);
                charItem.PowerRunScenario = ExcelOperation.GetCellValue(sheet, row, powerRunIndex);
                charItem.HarvFstp = ExcelOperation.GetCellValue(sheet, row, harvfstpIndex);
                charItem.SiteFlag = ExcelOperation.GetCellValue(sheet, row, siteflagIndex);
                charItem.FailFlag = ExcelOperation.GetCellValue(sheet, row, failflagIndex);
                charItem.ManualAc = ExcelOperation.GetCellValue(sheet, row, manualAcIndex);
                charItem.FailInfo = ExcelOperation.GetCellValue(sheet, row, failInfoIndex);
                charItem.Environment = ExcelOperation.GetCellValue(sheet, row, envIndex);
                charItem.Burst = ExcelOperation.GetCellValue(sheet, row, burstIndex);
                charItem.Wait = ExcelOperation.GetCellValue(sheet, row, waitIndex);
                charItem.TestInstanceName = ExcelOperation.GetCellValue(sheet, row, instanceNameIndex);
                charItem.Voltage = ExcelOperation.GetCellValue(sheet, row, voltageIndex);
                charItem.DcSelector = ExcelOperation.GetCellValue(sheet, row, dcSelectorIndex);
                charItem.AcSelector = ExcelOperation.GetCellValue(sheet, row, acSelectorIndex);
                charItem.AcCategory = ExcelOperation.GetCellValue(sheet, row, acCategoryIndex);
                charItem.DcCategory = ExcelOperation.GetCellValue(sheet, row, dcCategoryIndex);
                charItem.Levels = ExcelOperation.GetCellValue(sheet, row, levelsIndex);
                charItem.InitPatternsCell = ExcelOperation.GetCellValue(sheet, row, initPatternsIndex);
                charItem.PayloadPatternsCell = ExcelOperation.GetCellValue(sheet, row, payloadPatternsIndex);
                SetPayload1(charItem);
                //charItem.InitPatternList = GetPatternList(charItem.InitPatterns);
                //charItem.PayloadPatternList = GetPatternList(charItem.PayloadPatterns);
                // Check DigSrc setting for init patterns
                charItem.DigSrcAssignment = ExcelOperation.GetCellValue(sheet, row, srcAssignmentIndex);
                CheckDigSrcSetting(charItem);

                charItem.SearchMethod = ExcelOperation.GetCellValue(sheet, row, searchMethodIndex);
                charItem.SelSrmSendbit = ExcelOperation.GetCellValue(sheet, row, selSrmSendBitIndex);
                charItem.IsUseRtosCmd = ExcelOperation.GetCellValue(sheet, row, rtosCmdIndex).Equals("Y", StringComparison.OrdinalIgnoreCase);
                charItem.SuspendDatalog = ExcelOperation.GetCellValue(sheet, row, suspendDatalogIndex);
                #region Parse Shmoo setup
                var charShmooSetup = new ShmooSetup
                {
                    PlanShmooSetupName = ExcelOperation.GetCellValue(sheet, row, shmooSetupIndex).Replace(' ', '_'), // 20170125 add by JN   
                    Timeset = charItem.Timeset,
                    IsUseCmd = charItem.IsUseRtosCmd,
                    SearchMethod = charItem.SearchMethod,
                    SuspendDatalog = charItem.SuspendDatalog
                };

                _UpdateShmooPin(sheet, row, xsweep1Index, charShmooSetup, "X");  // xSweep1
                _UpdateShmooPin(sheet, row, xsweep1Index + 4, charShmooSetup, "X");  //xSweep2
                _UpdateShmooPin(sheet, row, xsweep1Index + 8, charShmooSetup, "X");  //xSweep3
                _UpdateShmooPin(sheet, row, ysweep1Index, charShmooSetup, "Y");  //ySweep1
                _UpdateShmooPin(sheet, row, ysweep1Index + 4, charShmooSetup, "Y");  //ySweep2
                _UpdateShmooPin(sheet, row, ysweep1Index + 8, charShmooSetup, "Y");  //ySweep3
                if (zsweep1Index != 0)
                {
                    _UpdateShmooPin(sheet, row, zsweep1Index, charShmooSetup, "Z");  //zSweep1
                    _UpdateShmooPin(sheet, row, zsweep1Index + 4, charShmooSetup, "Z");  //zSweep2
                    _UpdateShmooPin(sheet, row, zsweep1Index + 8, charShmooSetup, "Z");  //zSweep3
                }
                //Check shmoo setup
                if (charShmooSetup.ShmooPins.Count > 0)
                {
                    CheckShmooSetting(charShmooSetup);
                    charItem.CharShmooSetup = charShmooSetup;
                }
                #endregion

                string condition = ExcelOperation.GetCellValue(sheet, row, conditionIndex);
                charItem.CharCondition = condition != "" ? ConvertCharCondition(condition) : "";
                charItem.IpUse1 = ExcelOperation.GetCellValue(sheet, row, ipuse1Index);
                charItem.IpUse2 = ExcelOperation.GetCellValue(sheet, row, ipuse2Index);
                charItem.IpUse3 = ExcelOperation.GetCellValue(sheet, row, ipuse3Index);
                charItem.IpUse4 = ExcelOperation.GetCellValue(sheet, row, ipuse4Index);
                charItem.IpUse5 = ExcelOperation.GetCellValue(sheet, row, ipuse5Index);
                charItem.IpUse6 = ExcelOperation.GetCellValue(sheet, row, ipuse6Index);
                charItem.ProgramTestName1 = ExcelOperation.GetCellValue(sheet, row, programTestName1Index);
                charItem.ProgramTestName2 = ExcelOperation.GetCellValue(sheet, row, programTestName2Index);
                charItem.Ttr = ExcelOperation.GetCellValue(sheet, row, ttrIndex) == "1";
                charItem.Htol = ExcelOperation.GetCellValue(sheet, row, htolIndex) == "1";
                charItem.Use = ExcelOperation.GetCellValue(sheet, row, useIndex).Equals("Use", StringComparison.InvariantCultureIgnoreCase)
                    || ExcelOperation.GetCellValue(sheet, row, useIndex) == "";
                charItem.MeasType = DataConvertor.GetMeasType(charItem.TestInstanceName).ToUpper();
                charItem.IsNeedMask = ExcelOperation.GetCellValue(sheet, row, isPatternExistIndex);
                charItem.IsDateNeedReverse = !string.IsNullOrEmpty(ExcelOperation.GetCellValue(sheet, row, isDsscReverseIndex));
                //Use timeset in pattern list csv as high priority
                string csvTimeset = SearchInfo.GetTimeset(charItem.Payload1);
                //charItem.Timeset = csvTimeset != "" ? csvTimeset : ExcelOperation.GetCellValue(sheet, row, timesetIndex);
                charItem.Timeset = CompareTimsetBetweenCsvAndPlan(csvTimeset, ExcelOperation.GetCellValue(sheet, row, timesetIndex));
                charItem.Visit = false;
                charItem.Select = false;
                charItem.ManualACfromTimeset = ExcelOperation.GetCellValue(sheet, row, manualACfromTimesetIndex);
                charItem.ShiftFreq = ExcelOperation.GetCellValue(sheet, row, shiftFreqIndex);
                charItem.AcCategoryOri = charItem.AcCategory;
                charItem.BypassShmooHole = ExcelOperation.GetCellValue(sheet, row, bypassShmooHoleIndex);
                charItem.RetentionRamp = ExcelOperation.GetCellValue(sheet, row, retentionRampIndex);
                charItem.DigSrcBitSize = ExcelOperation.GetCellValue(sheet, row, digSrcBitSizeIndex);
                charItem.DigSrcSeg = ExcelOperation.GetCellValue(sheet, row, digSrcSegIndex);
                charItem.DigSrcPin = ExcelOperation.GetCellValue(sheet, row, digSrcPinIndex);
                charItem.DigSrcEq = ExcelOperation.GetCellValue(sheet, row, digSrcEqIndex);
                charItem.OneTimeInit = ExcelOperation.GetCellValue(sheet, row, oneTimeInitIndex);
                charItem.MappingPatternSet = ExcelOperation.GetCellValue(sheet, row, mappingPatternSetIndex);
                charItem.ApplyVoltageFromBinCut = ExcelOperation.GetCellValue(sheet, row, applyVoltageFromBinCutIndex);
                charItem.FreeRunningClock = ExcelOperation.GetCellValue(sheet, row, freeRunningClockIndex);
                charItem.IsFreeRunClk = !string.IsNullOrWhiteSpace(charItem.FreeRunningClock);
                charItem.UserFunction = ExcelOperation.GetCellValue(sheet, row, userFunctionIndex);
                HandlePatternsAfterSelsram(charItem);
                charItem.HarvestPinGrpOtherFail = ExcelOperation.GetCellValue(sheet, row, harvestPinGrpOtherFail);
                charItem.EnableCoreHarvest = ExcelOperation.GetCellValue(sheet, row, enableCoreHarvest);
                charItem.EnableCoreMask = ExcelOperation.GetCellValue(sheet, row, enableCoreMask);
                charItem.PinGrpSpecifyMask = ExcelOperation.GetCellValue(sheet, row, pinGrpSpecifyMask);
                charItem.SsnSpecifyMask = ExcelOperation.GetCellValue(sheet, row, ssnSpecifyMask);
                charItem.AdaptiveCooling = ExcelOperation.GetCellValue(sheet, row, adaptiveCooling);
                charItem.SelsramUserDef9 = ExcelOperation.GetCellValue(sheet, row, selsramUserDef9);
                charItem.StageCp1 = ExcelOperation.GetCellValue(sheet, row, stageCp1);
                charItem.StageCp2 = ExcelOperation.GetCellValue(sheet, row, stageCp2);
                charItem.StageFt1 = ExcelOperation.GetCellValue(sheet, row, stageFt1);
                charItem.StageFt2 = ExcelOperation.GetCellValue(sheet, row, stageFt2);
                charItem.Die = ExcelOperation.GetCellValue(sheet, row, die);
                UpdatePatternUsedInCharPlan(charItem.UsedPatterns.Where(p => p != "").ToList());
                newCharSheet.CharList.Add(charItem);
            }
            return newCharSheet;
        }

        private static void SetPayload1(CharPlanItem charItem)
        {
            if (charItem.PayloadPatternsDict.TryGetValue("PL1", out string value))
            {
                charItem.Payload1 = value;
                return;
            }

            if (charItem.PayloadPatternsDict.TryGetValue("PL2", out string value1))
            {
                charItem.Payload1 = value1;
            }
        }

        private static void HandlePatternsAfterSelsram(CharPlanItem charItem)
        {
            if (!LocalSpecs.InputParam.RunPayloadAfterSelsram || !LocalSpecs.InputParam.UseNewTChar)
            {
                return;
            }

            bool findSelsram = false;
            var newInitPatternsDict = new Dictionary<string, string>();
            var newPayloadPatternsDict = new Dictionary<string, string>();
            var convertIdxDict = new Dictionary<string, string>();
            foreach (KeyValuePair<string, string> initPattern in charItem.InitPatternsDict)
            {
                if (!findSelsram)
                {
                    if (initPattern.Value.IndexOf("SRMDSSC", StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        charItem.SelsramPatternIdx = initPattern.Key;
                        findSelsram = true;
                    }
                    newInitPatternsDict.Add(initPattern.Key, initPattern.Value);
                }
                else
                {
                    charItem.RunPayloadAfterSelsramShiftCount++;
                    string newInitIndex = "PL" + charItem.RunPayloadAfterSelsramShiftCount;
                    newPayloadPatternsDict.Add(newInitIndex, initPattern.Value);
                    convertIdxDict.Add(initPattern.Key, newInitIndex);
                }
            }
            charItem.InitPatternsDict = newInitPatternsDict;
            foreach (KeyValuePair<string, string> payloadPattern in charItem.PayloadPatternsDict)
            {
                string newPayloadIndex = "PL" + (int.Parse(payloadPattern.Key.Replace("PL", "")) + charItem.RunPayloadAfterSelsramShiftCount);
                newPayloadPatternsDict.Add(newPayloadIndex, payloadPattern.Value);
                convertIdxDict.Add(payloadPattern.Key, newPayloadIndex);
            }
            charItem.PayloadPatternsDict = newPayloadPatternsDict;

            charItem.PowerRunScenario = DataConvertor.ConvertPwrPatternIndex(convertIdxDict, charItem.PowerRunScenario);
            charItem.Wait = DataConvertor.ConvertPatternIndex(convertIdxDict, DataConvertor.ConvertWaitFromOldtoNew(charItem.Wait));
            charItem.RetentionRamp = DataConvertor.ConvertPatternIndex(convertIdxDict, charItem.RetentionRamp);
            charItem.DigSrcBitSize = DataConvertor.ConvertPatternIndex(convertIdxDict, charItem.DigSrcBitSize);
            charItem.DigSrcSeg = DataConvertor.ConvertPatternIndex(convertIdxDict, charItem.DigSrcSeg);
            charItem.DigSrcPin = DataConvertor.ConvertPatternIndex(convertIdxDict, charItem.DigSrcPin);
            charItem.DigSrcEq = DataConvertor.ConvertPatternIndex(convertIdxDict, charItem.DigSrcEq);
        }

        private static void _UpdateShmooPin(ExcelWorksheet sheet, int row, int colIndex, ShmooSetup charShmooSetup, string pinXy)
        {
            string sweep = ExcelOperation.GetCellValue(sheet, row, colIndex);
            if (sweep == "")
            {
                return;
            }

            string start = ExcelOperation.GetCellValue(sheet, row, colIndex + 1);
            string stop = ExcelOperation.GetCellValue(sheet, row, colIndex + 2);
            string step = ExcelOperation.GetCellValue(sheet, row, colIndex + 3);

            charShmooSetup.ShmooPins.Add(new ShmooPin(sweep, start, stop, step, pinXy));
        }

        private static void CheckShmooSetting(ShmooSetup charShmooSetup)
        {
            var existSetupList = LocalSpecs.AllShmooSetups.Where(a => a.PlanShmooSetupName == charShmooSetup.PlanShmooSetupName).ToList();
            if (existSetupList.Count == 0)
            {
                if (charShmooSetup.PlanShmooSetupName.Contains("SweepCode"))
                {
                    return;
                }

                charShmooSetup.ShmooSetupName = charShmooSetup.PlanShmooSetupName;
                LocalSpecs.AllShmooSetups.Add(charShmooSetup);
            }
            else
            {
                CompareShmooSetup(charShmooSetup, existSetupList);
            }
        }

        private static void CompareShmooSetup(ShmooSetup charShmooSetup, IEnumerable<ShmooSetup> existSetupList)
        {
            int maxIndex = 0;
            foreach (ShmooSetup existSetup in existSetupList)
            {
                string[] arrary = existSetup.ShmooSetupName.Split('_');
                string indexStr = arrary[arrary.Length - 1];
                if (Regex.IsMatch(indexStr, @"^\d+$"))
                {
                    int index = Convert.ToInt32(indexStr);
                    if (index > maxIndex)
                    {
                        maxIndex = index;
                    }
                }
                if (charShmooSetup.ShmooPins.Count != existSetup.ShmooPins.Count)
                {
                    continue;
                }

                if (existSetup.ShmooPins.All(shmooPin => charShmooSetup.ShmooPins.Exists(a => a.Equals(shmooPin))))
                {
                    charShmooSetup.ShmooSetupName = existSetup.ShmooSetupName;
                    return;
                }
            }
            maxIndex++;
            charShmooSetup.ShmooSetupName = charShmooSetup.PlanShmooSetupName + "_" + maxIndex;
            LocalSpecs.AllShmooSetups.Add(charShmooSetup);
        }

        private static void CheckDigSrcSetting(CharPlanItem charItem)
        {
            var patternFields = new List<string>()
            {
                "INIT1",
                "INIT2",
                "INIT3",
                "INIT4",
                "INIT5",
                "INIT6",
                "INIT7",
                "INIT8",
                "INIT9",
                "INIT10",
                "PL1",
                "PL2",
                "PL3",
                "PL4",
                "PL5",
            };
            if (charItem.DigSrcAssignment == "")
            {
                return;
            }

            var digSrc = new List<string>();
            string[] assignmentArray = charItem.DigSrcAssignment.Split(',');
            int patternIndex = 0;
            foreach (string assignment in assignmentArray)
            {
                if (assignment == "")
                {
                    digSrc.Add("");
                }
                else
                {
                    if (charItem.AllPatternsDict.ContainsKey(patternFields[patternIndex]))
                    {
                        if (!charItem.AllPatternsDict[patternFields[patternIndex]].Contains(","))
                        {
                            digSrc.Add(charItem.AllPatternsDict[patternFields[patternIndex]] + ":" + assignment);
                        }
                        else
                        {
                            foreach (string subPattern in charItem.AllPatternsDict[patternFields[patternIndex]].Split(','))
                            {
                                if (subPattern.Contains(":"))
                                {
                                    digSrc.Add(subPattern + ":" + assignment);
                                }
                                else
                                {
                                    digSrc.Add(subPattern);
                                }
                            }
                        }
                    }
                }
                patternIndex++;
                if (patternIndex > 14)
                {
                    break;
                }
            }
            charItem.DigSrc = string.Join(",", digSrc);
        }

        private static string ConvertCharCondition(string charCondition)
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

        private static void ReadPatSet_AI(ExcelWorksheet sheet)
        {
            var patSet = new PatSetSheet("PatSetMerge");
            Dictionary<string, int> headerOrder = ExcelOperation.GetHeaderOrder(sheet, 3); //valid order in row to start
            //	TD Group							
            int patternSetIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Pattern Set");
            int tdGroupIndex = ExcelOperation.GetHeaderIndex(headerOrder, "TD Group");
            int timeDomainIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Time Domain");
            int enableIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Enable");
            int fileGroupNameIndex = ExcelOperation.GetHeaderIndex(headerOrder, "File/Group Name");
            int burstIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Burst");
            int startLabelIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Start Label");
            int stopLabelIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Stop Label");
            int commentIndex = ExcelOperation.GetHeaderIndex(headerOrder, "Comment");


            for (int rowstartIndex = 4; rowstartIndex <= sheet.Dimension.End.Row; rowstartIndex++)
            {
                var row = new PatSet
                {
                    PatSetName = ExcelOperation.GetCellValue(sheet, rowstartIndex, patternSetIndex)
                };

                var patSetrow = new PatSetRow
                {
                    TdGroup = ExcelOperation.GetCellValue(sheet, rowstartIndex, tdGroupIndex),
                    TimeDomain = ExcelOperation.GetCellValue(sheet, rowstartIndex, timeDomainIndex),
                    Enable = ExcelOperation.GetCellValue(sheet, rowstartIndex, enableIndex),
                    File = ExcelOperation.GetCellValue(sheet, rowstartIndex, fileGroupNameIndex),
                    Burst = ExcelOperation.GetCellValue(sheet, rowstartIndex, burstIndex),
                    StartLabel = ExcelOperation.GetCellValue(sheet, rowstartIndex, startLabelIndex),
                    StopLabel = ExcelOperation.GetCellValue(sheet, rowstartIndex, stopLabelIndex),
                    Comment = ExcelOperation.GetCellValue(sheet, rowstartIndex, commentIndex),
                };
                row.AddRow(patSetrow);
                patSet.AddRow(row);
            }
            LocalSpecs.PatSetAll = patSet;
        }

        private static List<EmaMappingItem> ReadEmaMappingTalbe(ExcelWorksheet sheet)
        {
            var result = new List<EmaMappingItem>();
            EmaMappingItem emaItem = null;
            for (int row = 1; row <= sheet.Dimension.Rows; row++)
            {
                if (!string.IsNullOrEmpty(sheet.Cells[row, 1].Text))
                {
                    emaItem = new EmaMappingItem();
                    result.Add(emaItem);
                    emaItem.Pattern = sheet.Cells[row, 1].Text;
                    emaItem.CasesList = new List<string>();
                    for (int i = 3; i <= sheet.Dimension.Columns; i++)
                    {
                        string value = sheet.Cells[row, i].Text;
                        if (string.IsNullOrEmpty(value))
                        {
                            break;
                        }

                        emaItem.CasesList.Add(value);
                    }
                }
                else
                {
                    if (emaItem == null)
                    {
                        continue;
                    }

                    var emaSgmt = new EmaSubset(sheet.Cells[row, 2].Text.Replace("#", " "));
                    int colIndex = 3;
                    emaItem.ReferenceSets.Add(sheet.Cells[row, 2].Text, new List<EmaSubset>());
                    emaItem.ReferenceSets[sheet.Cells[row, 2].Text].Add(emaSgmt);
                    foreach (string caseInfo in emaItem.CasesList)
                    {
                        emaSgmt.Data.Add(caseInfo, sheet.Cells[row, colIndex].Text);
                        colIndex++;
                    }
                }
            }

            return result;
        }

        private static void UpdatePatternUsedInCharPlan(IEnumerable<string> patternsInPlanItem)
        {
            foreach (string pattern in patternsInPlanItem.Where(pattern =>
                !LocalSpecs.PatternsInCharPlan.Contains(pattern, StringComparer.InvariantCultureIgnoreCase)))
            {
                LocalSpecs.PatternsInCharPlan.Add(pattern);
            }
        }

        private static string CompareTimsetBetweenCsvAndPlan(string timeSetCsv, string timeSetPlan)
        {
            if (timeSetCsv.IndexOf(timeSetPlan, StringComparison.OrdinalIgnoreCase) != -1)
            {
                return timeSetCsv;
            }

            if (timeSetPlan.IndexOf(timeSetCsv, StringComparison.OrdinalIgnoreCase) != -1)
            {
                return timeSetPlan;
            }

            if (timeSetPlan != "")
            {
                return timeSetPlan;
            }

            return timeSetCsv;
        }
    }
}
