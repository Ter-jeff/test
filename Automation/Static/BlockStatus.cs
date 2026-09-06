using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.Scan.Harvest;

using CommonLib.Enums;

using TestPlanLib.Static;

namespace Automation.Static
{
    public class Status
    {
        public bool Down;
        public bool Enable;
    }

    public static class BlockStatus
    {
        //Automation
        public const string PreAction = "PreAction";
        public const string Scan = "Scan";
        public const string Mbist = "Mbist";
        public const string Basic = "Basic";
        public const string Evs = "EVS";
        public const string HardIp = "HardIP";
        public const string Efuse = "Efuse";
        public const string Rtos = "Rtos";
        public const string BinCut = "BinCut";
        public const string Htol = "HTOL";
        public const string Qa = "QA";
        public const string Ids = "IDS";

        //LCD
        public const string Lcd = "LCD";
        public const string Otp = "OTP";
        //RF
        public const string Dvdc = "DVDC";

        //Validation
        public const string TestLimit = "TestLimit";
        public const string Pat = "Pat";
        public const string DcTest = "DCTest";
        public const string BinOut = "BinOut";
        public const string PatternUsage = "PatternUsage";
        public const string MissingItem = "MissingItem";
        public const string TimeSetCheck = "TimesetCheck";
        public const string PinSetting = "PinSetting";
        public const string TestParameter = "TestParameter";

        private static readonly Dictionary<string, Status> _automationBlockStatus = new Dictionary<string, Status>();

        static BlockStatus()
        {
            Clear();
        }

        public static void Clear()
        {
            Create();
        }

        public static void Create()
        {
            _automationBlockStatus.Clear();

            _automationBlockStatus.Add(PreAction, new Status { Down = true, Enable = true });
            _automationBlockStatus.Add(Scan, new Status());
            _automationBlockStatus.Add(Mbist, new Status());
            _automationBlockStatus.Add(Basic, new Status());
            _automationBlockStatus.Add(Evs, new Status());
            _automationBlockStatus.Add(HardIp, new Status());
            _automationBlockStatus.Add(Efuse, new Status());
            _automationBlockStatus.Add(Rtos, new Status());
            _automationBlockStatus.Add(BinCut, new Status());
            _automationBlockStatus.Add(Htol, new Status());

            _automationBlockStatus.Add(Otp, new Status());
            _automationBlockStatus.Add(Dvdc, new Status());
            _automationBlockStatus.Add(Lcd, new Status());
            _automationBlockStatus.Add(Qa, new Status());
        }

        public static Status GetAutomationBlockStatus(string blockName)
        {
            if (_automationBlockStatus.TryGetValue(blockName, out Status status))
            {
                return status;
            }

            return null;
        }

        public static void SetAutomation(List<string> testPlans, List<string> scghList, string patternListCsvFile, List<string> binCuts, List<string> postBinCuts, string otp, string yaml)
        {
            bool hasHardip = false;
            bool hasHardipInScgh = false;
            bool hasEfuseTable = false;
            bool hasScan = false;
            bool hasMbist = false;
            bool hasSpi = false;
            bool hasPatternListCsv = false;
            bool hasPlanValidate = false;
            int igxlConsist = 0;
            bool hasIgxlConsist;
            bool hasBinCutFile = false;
            bool hasBinCutPostFile = false;
            bool hasOtpFile = false;
            bool hasDvdcFile = false;
            bool hasPlanTemplate = false;
            bool hasTestSetting = false;
            bool hasHtol = false;

            bool skipCsvDoc = true;
            bool skipScghDoc = true;

            if (testPlans.Any())
            {
                igxlConsist = 1;
                hasPlanValidate = true;
            }

            if (!string.IsNullOrEmpty(patternListCsvFile))
            {
                igxlConsist |= 1 << 1;
                hasPatternListCsv = true;
            }


            if (scghList.Any())
            {
                igxlConsist |= 1 << 2;
                hasPlanTemplate = true;
            }

            if (binCuts.Any())
            {
                hasBinCutFile = true;
            }

            if (postBinCuts.Any())
            {
                hasBinCutPostFile = true;
            }

            if (!string.IsNullOrEmpty(yaml))
            {
                hasOtpFile = true;
            }

            bool hasEvs = testPlans.Exists(x => x.StartsWith("Instance_EVS", StringComparison.CurrentCultureIgnoreCase));
            var sheets = testPlans.Where(x => x.Equals(NeededSheets.ArraysSize, StringComparison.CurrentCultureIgnoreCase) || x.Equals("EFUSE_BitDef_Table", StringComparison.CurrentCultureIgnoreCase)).ToList();
            bool hasEfuseSheets = sheets.Any() && sheets.Count == 2;
            bool hasEfuseInstanceSheets = testPlans.Exists(x => x.StartsWith("Instance_EFUSE", StringComparison.CurrentCultureIgnoreCase));
            hasSpi = testPlans.Exists(x => x.StartsWith("Rtos_", StringComparison.CurrentCultureIgnoreCase));

            if ((hasEfuseSheets && hasPatternListCsv && hasPlanTemplate) || (hasEfuseSheets && hasEfuseInstanceSheets))
            {
                hasEfuseTable = true;
            }

            ScanTestPlanSheets(testPlans, ref hasHardip, ref hasDvdcFile, ref hasTestSetting, ref hasOtpFile, ref hasHtol);

            if (ScanNonBinCutInstanceMain.GetBinCutInstanceSheets(testPlans).Any())
            {
                hasScan = true;
            }

            ScanScghSheets(scghList, ref hasHardipInScgh, ref hasScan, ref hasMbist, ref hasSpi);

            hasIgxlConsist = ComputeIgxlConsist(igxlConsist, skipCsvDoc, skipScghDoc);

            if (hasIgxlConsist)
            {
                ApplyAutomationStatus(hasEfuseTable, hasHardip, hasScan, hasMbist, hasSpi, hasBinCutFile, hasBinCutPostFile, hasOtpFile, hasDvdcFile, hasPlanValidate, hasTestSetting, hasHtol, hasEvs, hasHardipInScgh, skipScghDoc, otp);
            }
        }

        private static void ScanTestPlanSheets(List<string> testPlans, ref bool hasHardip, ref bool hasDvdcFile, ref bool hasTestSetting, ref bool hasOtpFile, ref bool hasHtol)
        {
            foreach (string sheet in testPlans)
            {
                if (sheet.StartsWith(NeededSheets.PrefixHardIp, StringComparison.OrdinalIgnoreCase) || sheet.StartsWith(NeededSheets.PrefixDctest, StringComparison.OrdinalIgnoreCase))
                {
                    hasHardip = true;
                }

                if (sheet.StartsWith(NeededSheets.PrefixWireless, StringComparison.OrdinalIgnoreCase) || sheet.StartsWith(NeededSheets.PrefixLcd, StringComparison.OrdinalIgnoreCase))
                {
                    hasDvdcFile = true;
                }

                if (NeededSheets.IsTestSettingSheetName(sheet, LocalSpecs.CurrentProject))
                {
                    hasTestSetting = true;
                }

                if (Regex.IsMatch(sheet, "ahb_register_map", RegexOptions.IgnoreCase))
                {
                    hasOtpFile = true;
                }

                if (sheet.StartsWith("Instance_Htol", StringComparison.OrdinalIgnoreCase))
                {
                    hasHtol = true;
                }
            }
        }

        private static void ScanScghSheets(List<string> scghList, ref bool hasHardipInScgh, ref bool hasScan, ref bool hasMbist, ref bool hasSpi)
        {
            foreach (string sheet in scghList)
            {
                if (LocalSpecs.Options.Device == EnumDevice.AP)
                {
                    ScanScghSheetAp(sheet, ref hasHardipInScgh, ref hasScan, ref hasMbist, ref hasSpi);
                }
                else if (LocalSpecs.Options.Device == EnumDevice.RF)
                {
                    ScanScghSheetRf(sheet, ref hasHardipInScgh, ref hasScan, ref hasMbist, ref hasSpi);
                }
            }
        }

        private static void ScanScghSheetAp(string sheet, ref bool hasHardipInScgh, ref bool hasScan, ref bool hasMbist, ref bool hasSpi)
        {
            if (Regex.IsMatch(sheet, NeededSheets.HardIpScghS + "|" + NeededSheets.HardIpScghC + "|" + NeededSheets.HardIpScghG + "|" + NeededSheets.HardIpScgh, RegexOptions.IgnoreCase))
            {
                hasHardipInScgh = true;
            }

            if (Regex.IsMatch(sheet, NeededSheets.ScanScghSoc + "|" + NeededSheets.ScanScghCpu + "|" + NeededSheets.ScanScghGpu, RegexOptions.IgnoreCase) || Regex.IsMatch(sheet, NeededSheets.RegexScanSheets, RegexOptions.IgnoreCase))
            {
                hasScan = true;
            }

            if (Regex.IsMatch(sheet, NeededSheets.MbistScghSoc + "|" + NeededSheets.MbistScghCpu + "|" + NeededSheets.MbistScghGpu, RegexOptions.IgnoreCase) || Regex.IsMatch(sheet, NeededSheets.RegexMbistSheets, RegexOptions.IgnoreCase) || Regex.IsMatch(sheet, NeededSheets.MbistCharScgSoc + "|" + NeededSheets.MbistCharScgCpu + "|" + NeededSheets.MbistCharScgGpu, RegexOptions.IgnoreCase) || Regex.IsMatch(sheet, NeededSheets.RegexMbistCharSheets, RegexOptions.IgnoreCase))
            {
                hasMbist = true;
            }

            if (Regex.IsMatch(sheet, NeededSheets.SpiScghCpu + "|" + NeededSheets.SpiScghGpu + "|" + NeededSheets.SpiScghSoc, RegexOptions.IgnoreCase))
            {
                hasSpi = true;
            }
        }

        private static void ScanScghSheetRf(string sheet, ref bool hasHardipInScgh, ref bool hasScan, ref bool hasMbist, ref bool hasSpi)
        {
            if (Regex.IsMatch(sheet, NeededSheets.HardIpScghS + "|" + NeededSheets.HardIpScghC + "|" + NeededSheets.HardIpScghG + "|" + NeededSheets.HardIpScgh, RegexOptions.IgnoreCase))
            {
                hasHardipInScgh = true;
            }

            if (Regex.IsMatch(sheet, NeededSheets.ScanScghSoc + "|" + NeededSheets.ScanScghCpu + "|" + NeededSheets.ScanScghGpu, RegexOptions.IgnoreCase))
            {
                hasScan = true;
            }

            if (Regex.IsMatch(sheet, NeededSheets.MbistScghSoc + "|" + NeededSheets.MbistScghCpu + "|" + NeededSheets.MbistScghGpu, RegexOptions.IgnoreCase))
            {
                hasMbist = true;
            }

            if (Regex.IsMatch(sheet, NeededSheets.SpiScghCpu + "|" + NeededSheets.SpiScghGpu + "|" + NeededSheets.SpiScghSoc, RegexOptions.IgnoreCase))
            {
                hasSpi = true;
            }
        }

        private static bool ComputeIgxlConsist(int igxlConsist, bool skipCsvDoc, bool skipScghDoc)
        {
            if (igxlConsist == 7)
            {
                return true;
            }
            else if (skipCsvDoc && igxlConsist == 5)
            {
                return true;
            }
            else if (skipScghDoc && (igxlConsist == 1 || igxlConsist == 3))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static void ApplyAutomationStatus(bool hasEfuseTable, bool hasHardip, bool hasScan, bool hasMbist, bool hasSpi, bool hasBinCutFile, bool hasBinCutPostFile, bool hasOtpFile, bool hasDvdcFile, bool hasPlanValidate, bool hasTestSetting, bool hasHtol, bool hasEvs, bool hasHardipInScgh, bool skipScghDoc, string otp)
        {
            if (LocalSpecs.Options.Device == EnumDevice.AP)
            {
                GetAutomationBlockStatus(Basic).Down = true;//hasTestSetting;
                GetAutomationBlockStatus(Efuse).Down = hasEfuseTable;
                GetAutomationBlockStatus(Scan).Down = hasScan;
                GetAutomationBlockStatus(Mbist).Down = hasMbist;
                GetAutomationBlockStatus(HardIp).Down = hasHardip;
                GetAutomationBlockStatus(BinCut).Down = hasBinCutFile || hasBinCutPostFile;
                GetAutomationBlockStatus(Evs).Down = hasEvs;
                GetAutomationBlockStatus(Rtos).Down = hasSpi;
                GetAutomationBlockStatus(Otp).Down = hasOtpFile;
                GetAutomationBlockStatus(Htol).Down = hasHtol;
                GetAutomationBlockStatus(Qa).Down = true;
            }
            else if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                GetAutomationBlockStatus(Basic).Down = hasPlanValidate;
                GetAutomationBlockStatus(otp).Down = hasOtpFile;
                GetAutomationBlockStatus(Lcd).Down = hasDvdcFile;
            }
            else if (LocalSpecs.Options.Device == EnumDevice.RF)
            {
                GetAutomationBlockStatus(Basic).Down = hasTestSetting;
                GetAutomationBlockStatus(Efuse).Down = hasEfuseTable;
                GetAutomationBlockStatus(BinCut).Down = false;
                GetAutomationBlockStatus(Evs).Down = hasEvs;
                GetAutomationBlockStatus(Rtos).Down = false;
                GetAutomationBlockStatus(Scan).Down = hasScan;
                GetAutomationBlockStatus(Mbist).Down = hasMbist;
                GetAutomationBlockStatus(HardIp).Down = hasHardip && (hasHardipInScgh || skipScghDoc);
                GetAutomationBlockStatus(Dvdc).Down = hasDvdcFile;
            }
        }
    }
}
