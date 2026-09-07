using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Cautogen.AutoCZ.CharPreProcessor.InputDataBase.CharPlan;
using Cautogen.AutoCZ.CharPreProcessor.InputReader.CharPlanReader;
using Cautogen.AutoCZ.CharPreProcessor.Utility;

using OfficeOpenXml;

namespace Cautogen.AutoCZ.CharPreProcessor.ImFile
{
    public class ImFileRow
    {
        public string ItemNum { get; set; }
        public string BlockName { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string TestInstName { get; set; }
        public string NamingSelection { get; set; }
        public string Voltage { get; set; }
        public string PowerRunScenario { get; set; }
        public string Wait { get; set; }
        public string ShmooSetupName { get; set; }
        public string XSweep1 { get; set; }
        public string XStart1 { get; set; }
        public string XStop1 { get; set; }
        public string XStep1 { get; set; }
        public string XSweep2 { get; set; }
        public string XStart2 { get; set; }
        public string XStop2 { get; set; }
        public string XStep2 { get; set; }
        public string XSweep3 { get; set; }
        public string XStart3 { get; set; }
        public string XStop3 { get; set; }
        public string XStep3 { get; set; }
        public string YSweep1 { get; set; }
        public string YStart1 { get; set; }
        public string YStop1 { get; set; }
        public string YStep1 { get; set; }
        public string YSweep2 { get; set; }
        public string YStart2 { get; set; }
        public string YStop2 { get; set; }
        public string YStep2 { get; set; }
        public string YSweep3 { get; set; }
        public string YStart3 { get; set; }
        public string YStop3 { get; set; }
        public string YStep3 { get; set; }
        public string ZSweep1 { get; set; }
        public string ZStart1 { get; set; }
        public string ZStop1 { get; set; }
        public string ZStep1 { get; set; }
        public string ZSweep2 { get; set; }
        public string ZStart2 { get; set; }
        public string ZStop2 { get; set; }
        public string ZStep2 { get; set; }
        public string ZSweep3 { get; set; }
        public string ZStart3 { get; set; }
        public string ZStop3 { get; set; }
        public string ZStep3 { get; set; }
        public string VbtFunction { get; set; }
        public string VbtParameterName { get; set; }
        public string VbtParameterValue { get; set; }
        public string DcCategory { get; set; }
        public string DcSelector { get; set; }
        public string AcCategory { get; set; }
        public string AcSelector { get; set; }
        public string Levels { get; set; }
        public string Timeset { get; set; }
        public string InitPatterns { get; set; }
        public string PayloadPatterns { get; set; }
        public string Retention { get; set; }
        public string IpUseTestItems { get; set; }
        public string IpUseDescription { get; set; }
        public string IpUseProgramTestName { get; set; }
        public string IpUseUsl { get; set; }
        public string IpUseLsl { get; set; }
        public string IpUseUnits { get; set; }
        public string Htol { get; set; }
        public string HardIp { get; set; }
        public string UseNotUse { get; set; }
        public string SearchMethod { get; set; }
        public string CharCondition { get; set; }
        public string Ttr { get; set; }
        public string ProgramTestName1 { get; set; }
        public string ProgramTestName2 { get; set; }
        public string DigSrcAssignment { get; set; }
        public string IsPatternNotExist { get; set; }
        public string IsRtosUseCmd { get; set; }
        public string SuspendDatalog { get; set; }
        public string HarvFstp { get; set; }
        public string SiteFlag { get; set; }
        public string ManualAc { get; set; }
        public string FailInfo { get; set; }
        public string Environment { get; set; }
        public string Burst { get; set; }
        public string IsDataReverse { get; set; }
        public string EnableWord { get; set; }
        public string FailFlag { get; set; }
        public string ManualACfromTimeset { get; set; }
        public string ShiftFreq { get; set; }
        public string BypassShmooHole { get; set; }
        public string RetentionRamp { get; set; }
        public string DigSrcBitSize { get; set; }
        public string DigSrcSeg { get; set; }
        public string DigSrcPin { get; set; }
        public string DigSrcEq { get; set; }
        public string OneTimeInit { get; set; }
        public string ApplyVoltageFromBinCut { get; set; }
        public string MappingPatternSet { get; set; }
        public string FreeRunningClock { get; set; }
        public string UserFunction { get; set; }
        public string HarvestPinGrpOtherFail { get; set; }
        public string EnableCoreHarvest { get; set; }
        public string EnableCoreMask { get; set; }
        public string PinGrpSpecifyMask { get; set; }
        public string SsnSpecifyMask { get; set; }
        public string AdaptiveCooling { get; set; }
        public string SelsrmUserDef9 { get; set; }
        public string StageCp1 { get; set; }
        public string StageCp2 { get; set; }
        public string StageFt1 { get; set; }
        public string StageFt2 { get; set; }
        public string Die { get; set; }
        public ImFileRow()
        {
        }

        private static readonly Dictionary<string, string> _acDictionary = new Dictionary<string, string>();
        private static int _acSerialNum = 1;
        public ImFileRow(Characterization charRow, string sheetName)
        {
            //20230725 take out for new mapping rule
            //var blockMap = InputDefReader.BlockMapDict.ContainsKey(charRow.Category + ":" + charRow.Group) ? InputDefReader.BlockMapDict[charRow.Category + ":" + charRow.Group] : null;
            //if (blockMap == null)
            //{
            //    blockMap = InputDefReader.BlockMapDict.ContainsKey(charRow.Category)
            //               ? InputDefReader.BlockMapDict[charRow.Category]
            //               : new InputDefRow();
            //}

            var shmooCondition = new SweepCondition(charRow);

            var extraInstanceNames = new List<string>();
            if (CharPlan.HardIpSheets.Contains(sheetName.ToUpper()) && !charRow.UserDef3.ToLower().Contains("mutli"))
            {
                string allInstanceNames = charRow.TpName.Replace(".", "p");
                UtilityMain.UtilityFunction.SubInstanceNames(allInstanceNames, extraInstanceNames);
            }

            ItemNum = "";
            BlockName = charRow.Category;
            Type = charRow.Group;
            TestInstName = charRow.TpName.Replace(".", "p");
            NamingSelection = ",,,,,,,,,,,,,,";
            Voltage = charRow.OtherSupplies.Split(':').First().Trim() != "multi" ? charRow.OtherSupplies.Split(':').First().Trim() : "NV";
            PowerRunScenario = _GetPowerRunScenario(charRow);
            Wait = charRow.Retention;
            ShmooSetupName = shmooCondition.SetupName;
            DcCategory = _GetDcCategoryLevel(charRow).Item1;
            DcSelector = charRow.DcSelector;
            AcCategory = _GetAcCategory(charRow);
            AcSelector = "Typ";
            Levels = _GetDcCategoryLevel(charRow).Item2;
            Timeset = _GetTimeset(charRow);
            IpUseProgramTestName = extraInstanceNames.Count > 0 ? extraInstanceNames[0] : "";
            InitPatterns = _GetInitPatterns(charRow);
            PayloadPatterns = _GetPayloadPatterns(charRow);
            Retention = charRow.Retention;
            IpUseTestItems = charRow.IpUse1;
            IpUseDescription = charRow.IpUse2;
            IpUseUsl = charRow.IpUse3;
            IpUseLsl = charRow.IpUse4;
            IpUseUnits = charRow.IpUse5;
            Htol = charRow.Htol;
            HardIp = CharPlan.HardIpSheets.Contains(sheetName.ToUpper()) ? "Y" : null;
            UseNotUse = charRow.Use;
            SearchMethod = charRow.Search;
            CharCondition = _GetInterPosePrePattern(charRow);
            Ttr = charRow.Ttr;
            DigSrcAssignment = charRow.DigSrc;
            ProgramTestName1 = extraInstanceNames.Count > 1 ? extraInstanceNames[1] : "";
            ProgramTestName2 = extraInstanceNames.Count > 2 ? extraInstanceNames[2] : "";

            XSweep1 = shmooCondition.XSweepList.Count >= 1 ? shmooCondition.XSweepList[0].SweepName : "";
            AssignXSweepFields(shmooCondition);

            AssignYSweepFields(shmooCondition);
            AssignZSweepFields(shmooCondition);

            AssignTrailingFields(charRow);

            ApplyDfcSuffix(charRow);
        }

        private void ApplyDfcSuffix(Characterization charRow)
        {
            if (!string.IsNullOrEmpty(charRow.Dfc))
            {
                if (!charRow.Dfc.Trim().Equals("0", StringComparison.OrdinalIgnoreCase))
                {
                    TestInstName = TestInstName + charRow.Dfc + "_";
                }
            }
        }

        private void AssignXSweepFields(SweepCondition shmooCondition)
        {
            XStart1 = shmooCondition.XSweepList.Count >= 1 ? shmooCondition.XSweepList[0].Start : "";
            XStop1 = shmooCondition.XSweepList.Count >= 1 ? shmooCondition.XSweepList[0].Stop : "";
            XStep1 = shmooCondition.XSweepList.Count >= 1 ? shmooCondition.XSweepList[0].Step : "";
            XSweep2 = shmooCondition.XSweepList.Count >= 2 ? shmooCondition.XSweepList[1].SweepName : "";
            XStart2 = shmooCondition.XSweepList.Count >= 2 ? shmooCondition.XSweepList[1].Start : "";
            XStop2 = shmooCondition.XSweepList.Count >= 2 ? shmooCondition.XSweepList[1].Stop : "";
            XStep2 = shmooCondition.XSweepList.Count >= 2 ? shmooCondition.XSweepList[1].Step : "";
            XSweep3 = shmooCondition.XSweepList.Count >= 3 ? shmooCondition.XSweepList[2].SweepName : "";
            XStart3 = shmooCondition.XSweepList.Count >= 3 ? shmooCondition.XSweepList[2].Start : "";
            XStop3 = shmooCondition.XSweepList.Count >= 3 ? shmooCondition.XSweepList[2].Stop : "";
            XStep3 = shmooCondition.XSweepList.Count >= 3 ? shmooCondition.XSweepList[2].Step : "";
        }

        private void AssignYSweepFields(SweepCondition shmooCondition)
        {
            YSweep1 = shmooCondition.YSweepList.Count >= 1 ? shmooCondition.YSweepList[0].SweepName : "";
            YStart1 = shmooCondition.YSweepList.Count >= 1 ? shmooCondition.YSweepList[0].Start : "";
            YStop1 = shmooCondition.YSweepList.Count >= 1 ? shmooCondition.YSweepList[0].Stop : "";
            YStep1 = shmooCondition.YSweepList.Count >= 1 ? shmooCondition.YSweepList[0].Step : "";
            YSweep2 = shmooCondition.YSweepList.Count >= 2 ? shmooCondition.YSweepList[1].SweepName : "";
            YStart2 = shmooCondition.YSweepList.Count >= 2 ? shmooCondition.YSweepList[1].Start : "";
            YStop2 = shmooCondition.YSweepList.Count >= 2 ? shmooCondition.YSweepList[1].Stop : "";
            YStep2 = shmooCondition.YSweepList.Count >= 2 ? shmooCondition.YSweepList[1].Step : "";
            YSweep3 = shmooCondition.YSweepList.Count >= 3 ? shmooCondition.YSweepList[2].SweepName : "";
            YStart3 = shmooCondition.YSweepList.Count >= 3 ? shmooCondition.YSweepList[2].Start : "";
            YStop3 = shmooCondition.YSweepList.Count >= 3 ? shmooCondition.YSweepList[2].Stop : "";
            YStep3 = shmooCondition.YSweepList.Count >= 3 ? shmooCondition.YSweepList[2].Step : "";
        }

        private void AssignZSweepFields(SweepCondition shmooCondition)
        {
            ZSweep1 = shmooCondition.ZSweepList.Count >= 1 ? shmooCondition.ZSweepList[0].SweepName : "";
            ZStart1 = shmooCondition.ZSweepList.Count >= 1 ? shmooCondition.ZSweepList[0].Start : "";
            ZStop1 = shmooCondition.ZSweepList.Count >= 1 ? shmooCondition.ZSweepList[0].Stop : "";
            ZStep1 = shmooCondition.ZSweepList.Count >= 1 ? shmooCondition.ZSweepList[0].Step : "";
            ZSweep2 = shmooCondition.ZSweepList.Count >= 2 ? shmooCondition.ZSweepList[1].SweepName : "";
            ZStart2 = shmooCondition.ZSweepList.Count >= 2 ? shmooCondition.ZSweepList[1].Start : "";
            ZStop2 = shmooCondition.ZSweepList.Count >= 2 ? shmooCondition.ZSweepList[1].Stop : "";
            ZStep2 = shmooCondition.ZSweepList.Count >= 2 ? shmooCondition.ZSweepList[1].Step : "";
            ZSweep3 = shmooCondition.ZSweepList.Count >= 3 ? shmooCondition.ZSweepList[2].SweepName : "";
            ZStart3 = shmooCondition.ZSweepList.Count >= 3 ? shmooCondition.ZSweepList[2].Start : "";
            ZStop3 = shmooCondition.ZSweepList.Count >= 3 ? shmooCondition.ZSweepList[2].Stop : "";
            ZStep3 = shmooCondition.ZSweepList.Count >= 3 ? shmooCondition.ZSweepList[2].Step : "";
        }

        private void AssignTrailingFields(Characterization charRow)
        {
            IsRtosUseCmd = charRow.IsUseRtosCmd ? "Y" : "";
            IsPatternNotExist = string.Join(",", charRow.IsPatternNotExist.Distinct());
            SuspendDatalog = charRow.SuspendDatalog;
            HarvFstp = charRow.HarvFstp;
            SiteFlag = charRow.SiteFlag;
            FailFlag = charRow.FailFlag;
            EnableWord = charRow.Enableword;
            ManualAc = _GetManualAC(charRow);
            FailInfo = charRow.FailInfo;
            Environment = charRow.Environment;
            Burst = charRow.Burst.Equals("burst", StringComparison.CurrentCultureIgnoreCase) ? "yes" : "";
            IsDataReverse = charRow.DigSrcShiftOrder ? "V" : "";
            ManualACfromTimeset = charRow.TimeSet.Split(':').Length > 1 ? charRow.TimeSet.Split(':')[1] : "";
            ShiftFreq = charRow.ShiftFreq;
            BypassShmooHole = charRow.BypassShmooHole == "1" || charRow.BypassShmooHole.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "";
            RetentionRamp = charRow.RetentionRamp;
            DigSrcBitSize = charRow.DigSrcBitSize;
            DigSrcSeg = charRow.DigSrcSeg;
            DigSrcPin = charRow.DigSrcPin;
            DigSrcEq = charRow.DigSrcEq;
            OneTimeInit = charRow.OneTimeInit.Equals("TRUE", StringComparison.CurrentCultureIgnoreCase) ? "TRUE" : "FALSE";
            ApplyVoltageFromBinCut = charRow.ApplyVoltageFromBinCut;
            MappingPatternSet = charRow.MappingSpec.PatternSet.FirstOrDefault();
            FreeRunningClock = charRow.FreeRunningClock;
            UserFunction = charRow.UserFunction;
            HarvestPinGrpOtherFail = charRow.HarvestPinGrpOtherFail;
            EnableCoreHarvest = charRow.EnableCoreHarvest;
            EnableCoreMask = charRow.EnableCoreMask;
            PinGrpSpecifyMask = charRow.PinGrpSpecifyMask;
            SsnSpecifyMask = charRow.SsnSpecifyMask;
            AdaptiveCooling = charRow.AdaptiveCooling;
            SelsrmUserDef9 = charRow.UserDef9;
            StageCp1 = charRow.StageCp1;
            StageCp2 = charRow.StageCp2;
            StageFt1 = charRow.StageFt1;
            StageFt2 = charRow.StageFt2;
            Die = charRow.Die;
        }

        public static void WriteHeader(ExcelWorksheet sh)
        {
            _WriteCell(sh, 1, "Item#");
            _WriteCell(sh, 2, "Block Name");
            _WriteCell(sh, 3, "Description");
            _WriteCell(sh, 4, "Type");
            _WriteCell(sh, 5, "Test Instance Name");
            _WriteCell(sh, 6, "Naming Selection");
            _WriteCell(sh, 7, "Voltage");
            _WriteCell(sh, 8, "Power Run Scenario");
            _WriteCell(sh, 9, "Wait");
            _WriteCell(sh, 10, "Shmoo Setup Name");
            _WriteCell(sh, 11, "X Sweep 1");
            _WriteCell(sh, 12, "X Start1");
            _WriteCell(sh, 13, "X Stop1");
            _WriteCell(sh, 14, "X Step1");
            _WriteCell(sh, 15, "X Sweep 2");
            _WriteCell(sh, 16, "X Start2");
            _WriteCell(sh, 17, "X Stop2");
            _WriteCell(sh, 18, "X Step2");
            _WriteCell(sh, 19, "X Sweep 3");
            _WriteCell(sh, 20, "X Start3");
            _WriteCell(sh, 21, "X Stop3");
            _WriteCell(sh, 22, "X Step3");
            _WriteCell(sh, 23, "Y Sweep 1");
            _WriteCell(sh, 24, "Y Start1");
            _WriteCell(sh, 25, "Y Stop1");
            _WriteCell(sh, 26, "Y Step1");
            _WriteCell(sh, 27, "Y Sweep 2");
            _WriteCell(sh, 28, "Y Start2");
            _WriteCell(sh, 29, "Y Stop2");
            _WriteCell(sh, 30, "Y Step2");
            _WriteCell(sh, 31, "Y Sweep 3");
            _WriteCell(sh, 32, "Y Start3");
            _WriteCell(sh, 33, "Y Stop3");
            _WriteCell(sh, 34, "Y Step3");
            _WriteCell(sh, 35, "Z Sweep 1");
            _WriteCell(sh, 36, "Z Start1");
            _WriteCell(sh, 37, "Z Stop1");
            _WriteCell(sh, 38, "Z Step1");
            _WriteCell(sh, 39, "Z Sweep 2");
            _WriteCell(sh, 40, "Z Start2");
            _WriteCell(sh, 41, "Z Stop2");
            _WriteCell(sh, 42, "Z Step2");
            _WriteCell(sh, 43, "Z Sweep 3");
            _WriteCell(sh, 44, "Z Start3");
            _WriteCell(sh, 45, "Z Stop3");
            _WriteCell(sh, 46, "Z Step3");
            _WriteCell(sh, 47, "VBT Function");
            _WriteCell(sh, 48, "VBT Parameter Name");
            _WriteCell(sh, 49, "VBT Parameter Value");
            _WriteCell(sh, 50, "DC Category");
            _WriteCell(sh, 51, "DC Selector");
            _WriteCell(sh, 52, "AC Category");
            _WriteCell(sh, 53, "AC Selector");
            _WriteCell(sh, 54, "Levels");
            _WriteCell(sh, 55, "Timeset");
            _WriteCell(sh, 56, "Init Patterns");
            _WriteCell(sh, 57, "Payload Patterns");
            _WriteCell(sh, 58, "Retention");
            _WriteCell(sh, 59, "IP Use", "Test Items");
            _WriteCell(sh, 60, "IP Use", "Description");
            _WriteCell(sh, 61, "IP Use", "ProgramTestName");
            _WriteCell(sh, 62, "IP Use", "USL");
            _WriteCell(sh, 63, "IP Use", "LSL");
            _WriteCell(sh, 64, "IP Use", "Units");
            _WriteCell(sh, 65, "HTOL");
            _WriteCell(sh, 66, "HardIP");
            _WriteCell(sh, 67, "Use/Not Use");
            _WriteCell(sh, 68, "Search Method");
            _WriteCell(sh, 69, "Char_Condition");
            _WriteCell(sh, 70, "TTR");
            _WriteCell(sh, 71, "ProgramTestName_1");
            _WriteCell(sh, 72, "ProgramTestName_2");
            _WriteCell(sh, 73, "DigSrcAssignment");
            _WriteCell(sh, 74, "PatternNotExist");
            _WriteCell(sh, 75, "RtosUseCmd");
            _WriteCell(sh, 76, "SuspendDatalog");
            _WriteCell(sh, 77, "Harv_FSTP");
            _WriteCell(sh, 78, "SiteFlag");
            _WriteCell(sh, 79, "Manual_AC");
            _WriteCell(sh, 80, "Fail_Info");
            _WriteCell(sh, 81, "Environment");
            _WriteCell(sh, 82, "Burst");
            _WriteCell(sh, 83, "IsEMA_Data_Reverse");
            _WriteCell(sh, 84, "Enable");
            _WriteCell(sh, 85, "FailFlag");
            _WriteCell(sh, 86, "ManualACfromTimeset");
            _WriteCell(sh, 87, "ShiftFreq");
            _WriteCell(sh, 88, "BypassShmooHole");
            _WriteCell(sh, 89, "RetentionRamp");
            _WriteCell(sh, 90, "DigSrcBitSize");
            _WriteCell(sh, 91, "DigSrcSeg");
            _WriteCell(sh, 92, "DigSrcPin");
            _WriteCell(sh, 93, "DigSrcEQ");
            _WriteCell(sh, 94, "OneTimeINIT");
            _WriteCell(sh, 95, "ApplyVoltageFromBinCut");
            _WriteCell(sh, 96, "MappingPatternSet");
            _WriteCell(sh, 97, "FreeRunningClock");
            _WriteCell(sh, 98, "UserFunction");
            _WriteCell(sh, 99, "HarvestPinGrpOtherFail");
            _WriteCell(sh, 100, "EnableCoreHarvest");
            _WriteCell(sh, 101, "EnableCoreMask");
            _WriteCell(sh, 102, "PinGrpSpecifyMask");
            _WriteCell(sh, 103, "SSNSpecifyMask");
            _WriteCell(sh, 104, "AdaptiveCooling");
            _WriteCell(sh, 105, "SelsrmUserDef9");
            _WriteCell(sh, 106, "StageCP1");
            _WriteCell(sh, 107, "StageCP2");
            _WriteCell(sh, 108, "StageFT1");
            _WriteCell(sh, 109, "StageFT2");
            _WriteCell(sh, 110, "DIE");
        }

        private static void _WriteCell(ExcelWorksheet sh, int col, string textRow1, string textRow2 = "")
        {
            sh.Cells[1, col].Value = textRow1;
            sh.Cells[1, col].Style.Font.Bold = true;
            sh.Cells[2, col].Value = textRow2;
            sh.Cells[2, col].Style.Font.Bold = true;
        }
        private static string _GetInitPatterns(Characterization charRow)
        {
            var result = new List<string>();
            foreach (PatternCell patternCell in charRow.PatternCellList)
            {
                if (patternCell.Header.StartsWith("init", StringComparison.OrdinalIgnoreCase))
                {
                    if (int.TryParse(Regex.Match(patternCell.Header, @"init(?<Num>\d+)", RegexOptions.IgnoreCase).Groups["Num"].ToString(), out int num))
                    {
                        result.Add($"INIT{num}={patternCell.PatternDefine.ToUpper()}");
                    }
                }
            }
            return string.Join(";", result);
        }
        private static string _GetPayloadPatterns(Characterization charRow)
        {
            var result = new List<string>();
            foreach (PatternCell patternCell in charRow.PatternCellList)
            {
                if (patternCell.Header.StartsWith("payload", StringComparison.OrdinalIgnoreCase))
                {
                    if (int.TryParse(Regex.Match(patternCell.Header, @"payload(?<Num>\d+)", RegexOptions.IgnoreCase).Groups["Num"].ToString(), out int num))
                    {
                        result.Add($"PL{num}={patternCell.PatternDefine.ToUpper()}");
                    }
                }
            }
            return string.Join(";", result);
        }
        private static Tuple<string, string> _GetDcCategoryLevel(Characterization charRow)
        {
            string dcCategory = "";
            string level = "";
            if (charRow.OtherSupplies.Split(' ').Length > 1)
            {
                dcCategory = charRow.OtherSupplies.Split(' ').First();
                level = charRow.OtherSupplies.Split(' ')[1].Split(':').Length > 1 ? charRow.OtherSupplies.Split(' ')[1].Split(':')[1].Trim() : _GetLevel(charRow, dcCategory);
            }
            if (string.IsNullOrEmpty(dcCategory))
            {
                if (charRow.MappingSpec.DcCategoryLevel.Any())
                {
                    var mappingDcLevelList = charRow.MappingSpec.DcCategoryLevel.ToList();
                    mappingDcLevelList = mappingDcLevelList
                    .OrderBy((str) =>
                    {
                        if (str.Split(':').Length > 1)
                        {
                            if (str.Split(':')[1].ToUpper().Contains(charRow.LevelsByBlock.ToUpper()))
                            {
                                if (str.Split(':')[1].ToUpper().Contains("EVS"))
                                {
                                    return 2;
                                }

                                return 1;
                            }
                            else
                            {
                                return 3;
                            }
                        }
                        return 4;
                    }).ToList();
                    dcCategory = mappingDcLevelList.First().Split(':')[0];
                    level = mappingDcLevelList.First().Split(':').Length > 1 ? mappingDcLevelList.First().Split(':')[1] : "";
                }
                else
                {
                    dcCategory = charRow.Category;
                    level = "";
                }
            }

            return Tuple.Create(dcCategory, level);
        }
        private static string _GetAcCategory(Characterization charRow)
        {
            string acCategory = "";
            if (charRow.TimeSet.Split(':').Length > 1)
            {
                acCategory = charRow.TimeSet.Split(':')[1];
                if (!string.IsNullOrWhiteSpace(charRow.ShiftFreq))
                {
                    acCategory += "_" + charRow.ShiftFreq + "MHz";
                }
            }
            else if (!string.IsNullOrEmpty(charRow.MappingSpec.FastestAcCategory))
            {
                List<string> namingList = charRow.MappingSpec.FastestAcCategory.Split('_').ToList();
                if (!string.IsNullOrWhiteSpace(charRow.ShiftFreq))
                {
                    if (namingList.Last().EndsWith("MHz", StringComparison.OrdinalIgnoreCase))
                    {
                        namingList.RemoveAt(namingList.Count - 1);
                    }
                    namingList.Add(charRow.ShiftFreq + "MHz");
                }
                acCategory = string.Join("_", namingList);
            }
            return acCategory;
        }

        private static string _GetLevel(Characterization charRow, string dc)
        {
            string mappingLevelsByDc = charRow.MappingSpec.DcCategoryLevel.FirstOrDefault(x => x.StartsWith(dc, StringComparison.OrdinalIgnoreCase));
            if (mappingLevelsByDc != null)
            {
                return mappingLevelsByDc.Split(':').Length > 1 ? mappingLevelsByDc.Split(':')[1] : "";
            }

            return charRow.MappingSpec.DcCategoryLevel.Any() ? charRow.MappingSpec.DcCategoryLevel.First().Split(':')[1].Trim() : "";
        }

        private static string _GetTimeset(Characterization charRow)
        {
            return !string.IsNullOrWhiteSpace(charRow.TimeSet.Split(':')[0]) ?
                            charRow.TimeSet.Split(':')[0] : charRow.MappingSpec.Timeset.Any() ?
                                                            charRow.MappingSpec.Timeset.First() : "";
        }

        private static string _GetPowerRunScenario(Characterization charRow)
        {
            return charRow.PowerRunScenario != "" ?
                                                    charRow.PowerRunScenario : "init_NV_pl_Sweep";
        }

        private static string _GetInterPosePrePattern(Characterization charRow)
        {
            string charCondition = string.Join(",", charRow.PowerSupplyX
                .Where(shmoo => !string.IsNullOrEmpty(shmoo.PowerCondition))
                .Select(shmoo =>
                {
                    string typeOrValt = shmoo.IsValt ? "Valt" : shmoo.Type;
                    typeOrValt = shmoo.IsVret ? "Vret" : shmoo.Type;
                    return $"{shmoo.Name}:{typeOrValt}:{shmoo.PowerCondition}";
                })
            ) + ",";

            charCondition = charRow.Pins.Where(pin => !string.IsNullOrEmpty(pin.PinCondition))
                .Aggregate(charCondition,
                    (current, pin) =>
                        current + pin.PinName.Replace(",", "+") + ":" + pin.PinType + ":" + pin.PinCondition + ",");

            string interposePrePattern = !string.IsNullOrEmpty(charRow.ApplyVoltageFromBinCut) ?
                charCondition.Replace(":VDD:", ":BV:") : charCondition.Replace(":VDD:", ":V:");

            if (charRow.IsFuncRow())
            {
                interposePrePattern += _GetLimitCondition(charRow);
            }

            // term or mask compare
            interposePrePattern = (from pin in charRow.Pins
                                   where !string.IsNullOrEmpty(pin.PinName)
                                   where pin.PinType.ToUpper() == "TERM" || pin.PinType.ToUpper() == "DISABLECOMPARE" || pin.PinType.ToUpper() == "ENABLECOMPARE"
                                   select pin)
                                   .Aggregate(interposePrePattern, (current, pin) => current + pin.PinName + ":" + pin.PinType + ";");

            return interposePrePattern.Replace(",", ";");
        }

        private static string _GetAcSymbolByPlan(string name)
        {
            string result = name;
            name = name.ToUpper();

            foreach (string symbol in UtilityMain.UtilityData.AcSpecsSymbols)
            {
                string symbolUpper = symbol.Trim().Replace("_", "").ToUpper();

                if (symbolUpper == name)
                {
                    return symbol.Trim().Replace("_", "");
                }

                if (Regex.Replace(symbolUpper, "VAR$", "") == name)
                {
                    return symbol.Trim().Replace("_", "");
                }

                if (Regex.Replace(symbolUpper, "FREQVAR$", "") == name)
                {
                    return symbol.Trim().Replace("_", "");
                }

                if (Regex.Replace(symbolUpper, "DIFFFREQVAR$", "") == name)
                {
                    return symbol.Trim().Replace("_", "");
                }
            }
            return result;
        }

        private static string _GetManualAC(Characterization charRow)
        {
            string charRowAc = _GetAcCategory(charRow);
            var tempList = new List<string>();
            string manualAc = "";
            charRow.PowerSupplyX.Where(x => x.PowerCondition == "" && x.Type != "VDD"
                                                                   && !string.IsNullOrEmpty(x.Start) && !string.IsNullOrEmpty(x.Stop))
                .ToList().ForEach(x =>
                {
                    string manualSymbol = _GetAcSymbolByPlan(x.Name.Split('(')[0]);
                    string manualSpeed = UtilityFunction.ConvertShiftSpeed(x.Start);
                    string manualAcCateName = Regex.IsMatch(charRow.AcStepName, "MHz", RegexOptions.IgnoreCase)
                                ? charRow.AcStepName
                                : charRow.AcStepName + "_" + UtilityFunction.ConvertShiftSpeed(x.Start).Replace(".", "p") + "MHz";
                    string originalAcCateName = charRowAc;
                    tempList.Add(string.Format($"{manualSymbol}:{manualSpeed}:{manualAcCateName}:{originalAcCateName}"));
                });
            manualAc = tempList.Any() ? string.Join(";", tempList) + ";" : "";

            //If more than 1 freq variable need to update
            if (!string.IsNullOrEmpty(manualAc))
            {
                string newManualAc = "";

                if (_acDictionary.TryGetValue(manualAc, out string value))  //same category that have appeared brfore, either 1 or more freq variable need to update
                {
                    foreach (string manual in manualAc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        newManualAc += manual.Split(':')[0] + ":" + manual.Split(':')[1] + ":" + value + ":" + charRowAc + ";";
                    }
                }
                //only 1 freq variable need to update, use original manual AC,but store the category in acDictionary
                else if (manualAc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Length == 1 && !_acDictionary.ContainsValue(manualAc.Split(';')[0].Split(':')[2]))
                {
                    string acDicValue = manualAc.Split(';')[0].Split(':')[2];
                    _acDictionary.Add(manualAc, acDicValue);
                    return manualAc;
                }
                else //more than 1 freq variable need to update, or with same category but different freq variable need to update, append serial number after orig category
                {
                    _acSerialNum = _acDictionary.Where(
                            x => Regex.IsMatch(x.Value, manualAc.Split(';')[0].Split(':')[2].Split('_')[0] + @"_\d+$", RegexOptions.IgnoreCase))
                        .ToList().Count + 1;

                    string acDicValue = manualAc.Split(';')[0].Split(':')[2].Split('_')[0] + "_" + _acSerialNum;
                    _acDictionary.Add(manualAc, acDicValue);

                    foreach (string manual in manualAc.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        newManualAc += manual.Split(':')[0] + ":" + manual.Split(':')[1] + ":" +
                                       _acDictionary[manualAc] + ":" + charRowAc + ";";

                    }
                }
                return newManualAc;
            }

            return manualAc;
        }

        private static string _GetLimitCondition(Characterization charRow)
        {
            /* return the limit portion of force condtion for input charRow */
            string usLcondition = "";
            if (Regex.IsMatch(charRow.IpUse3, @"\d"))
            {
                usLcondition = "USL:" + charRow.IpUse3;
            }

            if (Regex.IsMatch(charRow.IpUse4, @"\d"))
            {
                usLcondition += ";LSL:" + charRow.IpUse4;
            }

            return usLcondition.Trim(';');
        }
    }
}
