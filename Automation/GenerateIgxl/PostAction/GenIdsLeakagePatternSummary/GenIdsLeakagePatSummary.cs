using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;
using CommonLib.IdsLeakageCell;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.DataStruct;

namespace Automation.GenerateIgxl.PostAction.GenIdsLeakagePatternSummary
{
    public class GenIdsLeakagePatSummary
    {
        public void WorkFlow(Dictionary<string, PatternData> allPatterns, List<HardIpInfo> hardIpInfo, Dictionary<string, PatSetSheet> patSetSheets)
        {
            if (!hardIpInfo.Any())
            {
                Response.Report($"HardIP_PatInfo_{LocalSpecs.CurrentProject}.log not exist!!", EnumMessageLevel.Error);
                return;
            }

            if (!patSetSheets.ToList().Exists(x => x.Value.Name.Equals("PatSets_All", StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            var idsLeakagePats = allPatterns.Where(x => x.Key.ContainsIgnoreCase("IDS") || x.Key.ContainsIgnoreCase("HIZ") || x.Key.ContainsIgnoreCase("HIGHZ")).ToList();
            var patResultList = new List<PatternResult>();
            var patSheetList = patSetSheets.FirstOrDefault(x => x.Value.Name.Equals("PatSets_All", StringComparison.OrdinalIgnoreCase)).Value.Rows.Select(x => x.PatSetName).ToList();
            KeyValuePair<string, PatSetSheet> hardipPatSheet = patSetSheets.FirstOrDefault(x => x.Value.Name.Equals("PatSets_HardIP", StringComparison.OrdinalIgnoreCase));
            List<string> patSheetIDSList = hardipPatSheet.Key != null ? hardipPatSheet.Value.Rows.Where(row => row.PatSetName.
            StartsWith("IDS", StringComparison.OrdinalIgnoreCase)).Select(row => row.PatSetName).ToList() : new List<string>();
            foreach (KeyValuePair<string, PatternData> pat in idsLeakagePats)
            {
                HardIpInfo patInfo = hardIpInfo.Find(x => x.Payload.Equals(pat.Key, StringComparison.CurrentCultureIgnoreCase));
                if (patInfo != null && (!string.IsNullOrEmpty(patInfo.Xpins) || !string.IsNullOrEmpty(patInfo.UnusedIoPins)))
                {
                    if (!patSheetList.Exists(x => x.Equals(patInfo.Payload, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    var patResult = new PatternResult { Pattern = patInfo.Payload, Xpins = new List<string>() };
                    patResult.Xpins.AddRange(patInfo.Xpins.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries));
                    patResult.UnusedIoPins = new List<string>();
                    patResult.UnusedIoPins.AddRange(patInfo.UnusedIoPins.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries));
                    patResult.PatternVersion = pat.Value.PatternVersion.ToUpper();
                    patResultList.Add(patResult);
                }
            }

            foreach (string idsPatSet in patSheetIDSList)
            {
                string result = Regex.Replace(idsPatSet, @"^IDS_", "", RegexOptions.IgnoreCase);
                PatternResult patResult = patResultList.Where(x => x.Pattern.Equals(result)).ToList().FirstOrDefault();
                if (patResult != null)
                {
                    var patResultNew = new PatternResult { Pattern = idsPatSet };
                    patResultNew.Xpins = new List<string>(patResult.Xpins);
                    patResultNew.UnusedIoPins = new List<string>(patResult.UnusedIoPins);
                    patResultNew.PatternVersion = patResult.PatternVersion;
                    patResultList.Add(patResultNew);
                }

            }

            if (patResultList.Any())
            {
                PrintIdsLeakageInfo(patResultList);
            }
        }

        private void PrintIdsLeakageInfo(IEnumerable<PatternResult> patResultList)
        {

            string file = Path.Combine(FolderStructure.DirCommonSheets, "IdsLeakagePatternSummary.xlsx");
            var excel = new ExcelPackage(new FileInfo(file));
            ExcelWorksheet sheet = excel.Workbook.AddSheet("PreCondition");
            int currRow = 1;
            string title = "Pattern,PatternVersion,Xpins,UnusedIoPins,Comment";
            currRow = sheet.Cells[currRow, 1].PrintExcelRow(title.Split(','));
            var data = new List<List<string>>();
            var ignoreNwirePins = new List<string>();

            var ioIgnoreListSheet = new IoIgnoreList(
                EpWorkbook.TestPlanWorkbook.Worksheets.FirstOrDefault(p =>
                    Regex.IsMatch(p.Name, "IO_ignore_list", RegexOptions.IgnoreCase)));

            List<string> igrRegPins = (ioIgnoreListSheet.IoIgnorePinsSheets ?? new Dictionary<string, string>())
                .Where(x => x.Value.Equals("PreCondition", StringComparison.OrdinalIgnoreCase))
                .Select(kv => kv.Key).ToList();

            foreach (EnumEquipment equipment in TestPlanStatic.Equipments)
            {
                if (equipment == EnumEquipment.UltraFlex)
                {
                    ignoreNwirePins.AddRange(NwireSingleton.Instance().SettingInfo.NwirePins.Where(x => !string.IsNullOrEmpty(x.OutClk)).Select(x => $"{x.OutClk}_PA").ToList());
                    ignoreNwirePins.AddRange(NwireSingleton.Instance().SettingInfo.NwirePins.Where(x => !string.IsNullOrEmpty(x.OutClkDiff)).Select(x => $"{x.OutClkDiff}_PA").ToList());
                }
                else if (equipment == EnumEquipment.UltraFlexPlus)
                {
                    ignoreNwirePins.AddRange(NwireSingleton.Instance().SettingInfo.NwirePins.Where(x => !string.IsNullOrEmpty(x.OutClk)).Select(x => x.OutClk).ToList());
                    ignoreNwirePins.AddRange(NwireSingleton.Instance().SettingInfo.NwirePins.Where(x => !string.IsNullOrEmpty(x.OutClkDiff)).Select(x => x.OutClkDiff).ToList());
                }
            }
            ignoreNwirePins.AddRange(NwireSingleton.Instance().SettingInfo.NwirePins.Select(x => x.RefClk).ToList());

            foreach (PatternResult patternResult in patResultList)
            {
                string join = string.Join(",", patternResult.Xpins.Where(x =>
                    !ignoreNwirePins.Any(y => y.Equals(x, StringComparison.OrdinalIgnoreCase)) &&
                    !igrRegPins.Any(pat => Regex.IsMatch(x, pat.Replace("*", "(.+)?"), RegexOptions.IgnoreCase))
                    ));
                string unUsedPinStr = string.Join(",", patternResult.UnusedIoPins.Where(x =>
                    !ignoreNwirePins.Any(y => y.Equals(x, StringComparison.OrdinalIgnoreCase)) &&
                    !igrRegPins.Any(pat => Regex.IsMatch(x, pat.Replace("*", "(.+)?"), RegexOptions.IgnoreCase))
                    ));

                List<List<string>> datas = new IdsLeakageCellCheck().WorkFlow(patternResult, join, unUsedPinStr);
                if (datas.Any())
                {
                    data.AddRange(datas);
                }
            }
            sheet.Cells[currRow, 1].PrintExcelRowByList(data);
            sheet.View.FreezePanes(2, 1);
            sheet.MergeColumn(1);
            sheet.MergeColumn(2);
            if (!Directory.Exists(FolderStructure.DirCommonSheets))
            {
                Directory.CreateDirectory(FolderStructure.DirCommonSheets);
            }
            sheet.ExportSheet(Path.Combine(FolderStructure.DirCommonSheets, "PreCondition.txt"));
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "PreCondition");
        }
    }
}
