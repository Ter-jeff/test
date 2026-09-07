using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using BinCutScriptLib.Base;
using BinCutScriptLib.Reader;
using BinCutScriptLib.SetFunction.SetStartStep;
using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using OfficeOpenXml;

using TestPlanLib;
using TestPlanLib.BinCut.Binning;

namespace BinCutScriptLib
{
    internal static class BinCutReadMainHelpers1
    {
        public static Dictionary<string, string> GetAllPerformanceModeDic(List<string> performanceModeList)
        {
            var performanceModeDic = new Dictionary<string, string>();
            int dig456 = 3;
            foreach (IGrouping<string, string> group in performanceModeList.GroupBy(x => x[..3]))
            {
                if (group.Count() >= 10)
                {
                    int mod = (group.Count() - 6) % dig456;
                    int step = (group.Count() - 6) / dig456;
                    var maxCnt = new List<int>();
                    for (int i = 0; i < dig456; i++)
                    {
                        maxCnt.Add(i < mod ? step + 1 : step);
                    }

                    int final = 0;
                    for (int index = 0; index < group.ToList().Count; index++)
                    {
                        int cnt = index + 1;
                        string mode = group.ToList()[index].ToUpper();
                        if (Reg.RegContainPerformanceModeWithGroup.IsMatch(mode))
                        {
                            if (cnt <= dig456)
                            {
                                if (!performanceModeDic.ContainsKey(mode))
                                {
                                    performanceModeDic.Add(mode, cnt.ToString());
                                }
                            }
                            else if (cnt > dig456 && cnt <= group.Count() - dig456)
                            {
                                int num = 0;
                                int total = 0;
                                for (int i = 0; i < maxCnt.Count; i++)
                                {
                                    total += maxCnt[i];
                                    if (total >= cnt - dig456)
                                    {
                                        num = i;
                                        break;
                                    }
                                }

                                if (!performanceModeDic.ContainsKey(mode))
                                {
                                    performanceModeDic.Add(mode, (num + 4).ToString());
                                }
                            }
                            else if (cnt > group.Count() - dig456)
                            {

                                if (!performanceModeDic.ContainsKey(mode))
                                {
                                    performanceModeDic.Add(mode, (final + 7).ToString());
                                    final++;
                                }
                            }
                        }
                    }
                }
                else
                {
                    bool flag = false;
                    foreach (string mode in group)
                    {
                        if (Reg.RegContainPerformanceModeWithGroup.IsMatch(mode))
                        {
                            string number = Reg.RegContainPerformanceModeWithGroup.Match(mode).Groups["modenumber"].ToString();
                            if (!int.TryParse(number, out int _))
                            {
                                flag = true;
                                break;
                            }
                        }
                    }

                    if (flag)
                    {
                        int cnt = 1;
                        foreach (string mode in group)
                        {
                            if (!performanceModeDic.ContainsKey(mode.ToUpper()))
                            {
                                performanceModeDic.Add(mode.ToUpper(), cnt.ToString());
                                cnt++;
                            }
                        }
                    }
                    else
                    {
                        foreach (string mode in group)
                        {
                            if (Reg.RegContainPerformanceModeWithGroup.IsMatch(mode))
                            {
                                string number = Reg.RegContainPerformanceModeWithGroup.Match(mode).Groups["modenumber"].ToString();
                                if (!performanceModeDic.ContainsKey(mode.ToUpper()))
                                {
                                    performanceModeDic.Add(mode.ToUpper(), number.Last().ToString());
                                }
                            }
                        }
                    }
                }
            }
            return performanceModeDic;
        }

        internal static bool CheckFlagNameIsExist(string flagItems, string flagName)
        {
            if (string.IsNullOrEmpty(flagItems))
            {
                return true;
            }

            string[] spt = flagItems.Split([',']);
            if (spt.Any(s => s.EqualsIgnoreCase(flagName)))
            {
                return true;
            }

            return false;
        }

        internal static void GetDcSpecs(InstanceSheet instanceSheet, List<Tuple<string, string>> dcCategorys, EnumJob enumJob, string tempFolder)
        {
            ExcelWorksheet? dcSpecSheet = null;
            if (BinCutData.JoblistSheet != null)
            {
                string dcspectxt = "";
                foreach (JobRow row in BinCutData.JoblistSheet.Rows)
                {
                    string name = new Job(row.JobName).JobType.ToString();
                    if (name.EqualsIgnoreCase(enumJob.ToString()))
                    {
                        dcspectxt = row.DcSpecs.Split(',').First();
                        break;
                    }
                }
                if (!string.IsNullOrEmpty(dcspectxt))
                {
                    dcSpecSheet = GetSheetFromTempFolder(dcspectxt + ".txt", tempFolder).First();
                }
            }
            DcSpecSheet? dcSpec = null;
            if (dcSpecSheet != null)
            {
                dcSpec = new ReadDcSpecSheet().ReadSheet(dcSpecSheet);
            }

            GlobalSpecSheet globalSpecSheet = IgxlSheetReaderHelpers.GetIgxlSheets(tempFolder, EnumSheetType.DTGlobalSpecSheet).OfType<GlobalSpecSheet>().First();
            var tableGenertor = new TableGenertor();

            tableGenertor.CreateTable(tempFolder, BinCutData.PowerPins, instanceSheet, dcCategorys, dcSpec!, globalSpecSheet, out VoltageTable vrsTable, out VoltageTable nvTable);
            BinCutData.VrsTable = vrsTable;
            BinCutData.NvTable = nvTable;

            tableGenertor.CreateTableCs(tempFolder, BinCutData.PowerPins, instanceSheet, dcCategorys, dcSpec!, globalSpecSheet, out VoltageTable vrsAllTable, out VoltageTable nvAllTable);
            BinCutData.VrsAllTable = vrsAllTable;
            BinCutData.NvAllTable = nvAllTable;
        }

        internal static string FindPwrBinningSheetName(string searchSheet, string tempFolder, Job job)
        {
            var pwrBinningSheets = new List<string>();
            string pwrBinningSheet = "";
            List<string> fileList = [.. Directory.GetFiles(tempFolder)];
            foreach (string file in fileList)
            {
                string sheet = Path.GetFileName(file);
                if (sheet.StartsWith(searchSheet))
                {
                    pwrBinningSheets.Add(sheet);
                }
            }
            if (pwrBinningSheets.Count > 1)
            {
                foreach (string sheetName in pwrBinningSheets)
                {
                    if (sheetName.Contains(job.JobName))
                    {
                        pwrBinningSheet = sheetName;
                    }
                }
            }

            if (pwrBinningSheet.Length == 0 && pwrBinningSheets.Count != 0)
            {
                pwrBinningSheet = pwrBinningSheets.First();
            }

            return pwrBinningSheet;
        }

        internal static List<ExcelWorksheet> GetSheetFromTempFolder(string fileName, string tempFolder)
        {
            List<string> fileList = [.. Directory.GetFiles(tempFolder, fileName)];
            if (fileList.Count > 0)
            {
                var sheetList = new List<ExcelWorksheet>();
                foreach (string file in fileList)
                {
                    var binCutExcelPackage = new ExcelPackage();
                    string newSheetName = Path.GetFileNameWithoutExtension(file);
                    ExcelWorksheet sheet = binCutExcelPackage.Workbook.Worksheets.Add(newSheetName);
                    int index = 0;
                    using (var sr = new StreamReader(file))
                    {
                        while (!sr.EndOfStream)
                        {
                            string? line = sr.ReadLine();
                            index++;
                            if (line != null)
                            {
                                string[] lineSpt = line.Split(['\t'], StringSplitOptions.None);
                                sheet.Cells[index, 1].PrintExcelRow(lineSpt);
                            }
                        }
                    }
                    sheetList.Add(sheet);
                }
                return sheetList;
            }
            return [];
        }

        internal static BinningTables GetVddBinningDefAndOtherRail(string tempFolder, Action<string, Color> richTextBoxAppend, Job job)
        {
            richTextBoxAppend("Reading Binning Sheets ...", Color.Blue);
            List<string> sVddBinList = [.. Directory.GetFiles(tempFolder, BinCutConfig.VddBinningDefAppA1 + ".txt")];
            if (sVddBinList.Count == 0)
            {
                sVddBinList = [.. Directory.GetFiles(tempFolder, BinCutConfig.BincutEqnAppA + ".txt")];
            }

            string[] sheet2 = Directory.GetFiles(tempFolder, BinCutConfig.VddBinningDefAppA2 + ".txt");
            string[] sheet3 = Directory.GetFiles(tempFolder, BinCutConfig.VddBinningDefAppA3 + ".txt");
            #region  overwrite vddbining with jobName if vddbinning use default value
            if (BinCutConfig.VddBinningDefAppA1 == "Vdd_Binning_Def_appA_1" &&
                File.Exists(Path.Combine(tempFolder, BinCutConfig.VddBinningDefAppA1 + "_" + job.JobName + ".txt")))
            {
                sVddBinList = [.. Directory.GetFiles(tempFolder, BinCutConfig.VddBinningDefAppA1 + "_" + job.JobName + ".txt")];
            }
            else if (BinCutConfig.BincutEqnAppA == "bincut_eqn_appA" &&
                File.Exists(Path.Combine(tempFolder, BinCutConfig.BincutEqnAppA + "_" + job.JobName + ".txt")))
            {
                sVddBinList = [.. Directory.GetFiles(tempFolder, BinCutConfig.BincutEqnAppA + "_" + job.JobName + ".txt")];
            }
            if (BinCutConfig.VddBinningDefAppA2 == "Vdd_Binning_Def_appA_2" &&
                File.Exists(Path.Combine(tempFolder, BinCutConfig.VddBinningDefAppA2 + "_" + job.JobName + ".txt")))
            {
                sheet2 = Directory
                    .GetFiles(tempFolder, BinCutConfig.VddBinningDefAppA2 + "_" + job.JobName + ".txt");
            }
            if (BinCutConfig.VddBinningDefAppA3 == "Vdd_Binning_Def_appA_3" &&
                File.Exists(Path.Combine(tempFolder, BinCutConfig.VddBinningDefAppA3 + "_" + job.JobName + ".txt")))
            {
                sheet3 = Directory
                    .GetFiles(tempFolder, BinCutConfig.VddBinningDefAppA3 + "_" + job.JobName + ".txt");
            }

            #endregion
            sVddBinList.AddRange(sheet2);
            sVddBinList.AddRange(sheet3);
            var binningTables = new BinningTables();
            for (int cnt = 0; cnt < sVddBinList.Count; cnt++)
            {
                BinningTable binningTable = BinningTableReader.Read(sVddBinList[cnt]);
                binningTables.Add(binningTable);
            }
            BinCutData.AllPowers = [.. binningTables.First().GetPowerNames().Distinct()];
            BinCutData.PinInfos = binningTables.First().GetModeVsPin();
            BinCutData.BinningTables = binningTables.GetBinningRailTables();
            BinCutData.OtherRailTables = binningTables.GetOtherRailTables();

            List<InterpolatioNode> newInterpolationList = BinCutReadMainHelpers.GetInterpolationListDic(BinCutData.BinningTables);
            if (newInterpolationList != null && newInterpolationList.Count != 0)
            {
                BinCutConfig.InterpolationNodes = newInterpolationList;
            }

            return binningTables;
        }

    }
}
