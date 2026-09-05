using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using Automation.GenerateIgxl.Wireless.DVDC.InputReader;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Static;

namespace Automation.GenerateIgxl.HardIp.InputReader
{
    public class TestPlanConverter
    {
        protected readonly ScghData ScghData;

        public TestPlanConverter(ScghData scghData)
        {
            ScghData = scghData;
        }

        private static readonly Regex _regWorkSheetKey = new Regex(Pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _hardIpRegex = new Regex("^HardIP_", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex _idsRegex = new Regex(SheetConst.DcTestIds, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private const string Pattern = "^HARDIP_|DCTEST_IDS|DCTEST_GPIO|DCTEST_FailSafe|PLLDEBUG_";


        public Dictionary<string, HardIpSheet> ReadHardipPatterns(ExcelWorkbook excelWorkbook, Func<string, bool> isValidSheet, List<string> usedPatterns = null)
        {
            var workSheets = new List<ExcelWorksheet>();
            foreach (ExcelWorksheet worksheet in excelWorkbook.Worksheets)
            {
                string sheetName = worksheet.Name;
                bool match = isValidSheet?.Invoke(sheetName) ?? _regWorkSheetKey.IsMatch(sheetName);

                if (!match)
                {
                    continue;
                }

                if (_hardIpRegex.IsMatch(sheetName) && !LocalSpecs.IsModuleIncluded(BlockConst.HardIp))
                {
                    Response.Report($"Here is one HardIP sheet :\"{sheetName}\", but HardIP isn't used in Flow_Main, it will be ignored!",
                        EnumMessageLevel.Warning, 30);
                    continue;
                }

                if (_idsRegex.IsMatch(sheetName) && !LocalSpecs.IsModuleIncluded("IDS"))
                {
                    Response.Report($"Here is one IDS sheet :\"{sheetName}\", but IDS isn't used in Flow_Main, it will be ignored!",
                        EnumMessageLevel.Warning, 30);
                    continue;
                }

                if (HardIpConstData.IgnoredHardIpSheetList.Any(s => s.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    Response.Report($"Here is one HardIP sheet :\"{sheetName}\" will be ignored!", EnumMessageLevel.Warning,
                        30);
                    continue;
                }
                workSheets.Add(worksheet);
            }

            Dictionary<string, HardIpSheet> dic = ReadSheet(workSheets, usedPatterns);

            return dic;
        }

        internal List<HardIpPattern> ParseTrimItems(ref List<HardIpPattern> patterns)
        {
            var trimPatterns = new List<HardIpPattern>();
            var postBurnPatterns = new List<HardIpPattern>();
            foreach (HardIpPattern pattern in patterns)
            {
                if (!string.IsNullOrEmpty(pattern.WirelessData.TrimTarget))
                {

                    HardIpPattern trimPattern = pattern.Copy();
                    trimPattern.MiscInfo += ";Run_NV_Flow_Only;";
                    trimPatterns.Add(trimPattern);
                    if (pattern.MiscInfo.ContainsIgnoreCase("postburn"))
                    {
                        pattern.WirelessData.IsNeedPostBurn = true;
                    }
                }
                else
                {
                    pattern.WirelessData.IsDoMeasure = true;
                }


                postBurnPatterns.Add(pattern);
            }
            patterns = postBurnPatterns;
            return trimPatterns;
        }

        public Dictionary<string, HardIpSheet> ReadSheet(List<ExcelWorksheet> workSheets, List<string> usedPatterns)
        {
            double count = workSheets.Count;
            int index = 0;
            var planDic = new ConcurrentDictionary<string, HardIpSheet>();
            ScghData scghData = ScghStatic.ScghData;
            //foreach (ExcelWorksheet worksheet in workSheets)
            Parallel.ForEach(workSheets, worksheet =>
            {
                string sheetName = worksheet.Name;
                int currentIndex = Interlocked.Increment(ref index);
                double p = 100.0 * currentIndex / count;

                Response.Report($"Extracting sheet: {sheetName} ...", percentage: Convert.ToInt32(p));
                TestPlanReader reader = (sheetName.StartsWith(NeededSheets.PrefixWireless, StringComparison.OrdinalIgnoreCase) ||
                    sheetName.StartsWith(NeededSheets.PrefixLcd, StringComparison.OrdinalIgnoreCase)) ?
                    new WirelessTestPlanReader() : new TestPlanReader();
                //var reader = new TestPlanReader();
                TestPlanSheet planSheet = reader.ReadSheet(worksheet);
                if (usedPatterns != null)
                {
                    planSheet.PatternRows = reader.FilterTestItemsByPatterns(planSheet.PatternRows, usedPatterns);
                }

                planSheet.ConvertRealPatternName(scghData);
                planSheet.DividePatternRow();
                ParsePlanSheet(planSheet);

                List<HardIpPattern> testPlanPatterns = planSheet.PatternItems;
                if (LocalSpecs.Options.Device == EnumDevice.LCD)
                {
                    ParseTrimItems(ref testPlanPatterns);
                }

                planSheet.UpdateRelay();
                var hardIpSheet = new HardIpSheet
                {
                    SheetName = planSheet.SheetName,
                    Rows = testPlanPatterns,
                    PlanHeaderIdx = planSheet.PlanHeaderIdx
                };
                planDic.TryAdd(sheetName, hardIpSheet);
            }
            );

            new PatternBurstResolver().SetPatternBurst(planDic);

            var dic = planDic.OrderBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            #region Add Cap-Src manager for register assignment. Added on 2016/4/21
            var regAssignParser = new RegisterAssignParser();
            regAssignParser.ParseRegisterAssign(dic);
            #endregion

            return dic;
        }

        protected void ParsePlanSheet(TestPlanSheet planSheet)
        {
            var tpPreProcess = new TestPlanSheetPatPreprocess(planSheet);
            tpPreProcess.UpdateSheetPattern();

            var testPlanPatParser = new TestPlanPatParser(planSheet);
            testPlanPatParser.ConvertTpPatterns();

            if (planSheet.SheetName.StartsWith(NeededSheets.PrefixWireless, StringComparison.OrdinalIgnoreCase) ||
                planSheet.SheetName.StartsWith(NeededSheets.PrefixLcd, StringComparison.OrdinalIgnoreCase))
            {
                testPlanPatParser.ConvertWirelessTpPatterns();
            }
        }
    }
}
