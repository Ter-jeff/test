using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using BinCutScriptLib.Static;

using CommonLib.Extension;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlReader;
using IgxlLib.IgxlSheets;
using IgxlLib.Utility;

using OfficeOpenXml;

namespace BinCutScriptLib.Reader
{
    internal class TestProgramReaderHelpers
    {
        private static List<BinTableRow> LoadBinTableRows(string testProgramFilePath, Action<string, Color> richTextBoxAppend)
        {
            var bintableList = new List<ExcelWorksheet>();
            ExcelWorksheet? excelWorksheet = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "Bin_Table");
            if (excelWorksheet != null)
            {
                bintableList.Add(excelWorksheet);
            }

            ExcelWorksheet? binTableBinCut = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "Bin_Table_BinCut");
            if (binTableBinCut != null)
            {
                bintableList.Add(binTableBinCut);
            }

            var binTableRows = new List<BinTableRow>();
            richTextBoxAppend("Reading BinTable ...", Color.Blue);
            foreach (ExcelWorksheet bintable in bintableList)
            {
                var binSheetReader = new ReadBinTableSheet();
                BinTableSheet binSheet = binSheetReader.ReadSheet(bintable);
                IEnumerable<BinTableRow> rows = binSheet.Rows.Where(x => !x.IsBackup);
                binTableRows.AddRange(rows);
            }

            return binTableRows;
        }

        private static void CheckBinCutFlags(List<BinTableRow> binTableRows, Action<string, Color> richTextBoxAppend)
        {
            var dic = new Dictionary<string, string>();
            if (BinCutConfig.VddbinningFailStopFlag.Length != 0)
            {
                dic.Add("Vddbinning_Fail_Stop", BinCutConfig.VddbinningFailStopFlag);
            }

            if (BinCutConfig.VddbinningPowerBinningFailStopFlag.Length != 0)
            {
                dic.Add("Vddbinning_Power_Binning_Fail_Stop", BinCutConfig.VddbinningPowerBinningFailStopFlag);
            }

            if (BinCutConfig.VddbinningIdsFailFlag.Length != 0)
            {
                dic.Add("Bin_Vddbinning_IDS_fail", BinCutConfig.VddbinningIdsFailFlag);
            }

            if (BinCutConfig.VddbinningInterpolationFailFlag.Length != 0)
            {
                dic.Add("Vddbinning_Interpolation_fail", BinCutConfig.VddbinningInterpolationFailFlag);
            }
            if (binTableRows.Count == 0 || dic.Count == 0)
            {
                return;
            }
            foreach (KeyValuePair<string, string> item in dic)
            {
                BinTableRow? binTable = binTableRows
                    .Find(x => x.Name.EqualsIgnoreCase(item.Key));
                if (binTable != null)
                {
                    if (!BinCutReadMainHelpers1.CheckFlagNameIsExist(binTable.ItemList, item.Value))
                    {
                        richTextBoxAppend($"{item.Value} is missmatch", Color.Red);
                    }
                }
                else
                {
                    richTextBoxAppend($"{item.Key} is missing", Color.Red);
                }
            }
        }

        private static void ReadInstanceSheet(string testProgramFilePath, bool isCsharp, Action<string, Color> richTextBoxAppend)
        {
            InstanceSheet? instSheet = null;
            ExcelWorksheet? testInstSheet = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Vddbinning");
            if (testInstSheet != null)
            {
                richTextBoxAppend("Reading TestInst_Vddbinning ...", Color.Blue);
                instSheet = new ReadInstanceSheet().ReadSheet(testInstSheet);
            }
            ExcelWorksheet? testInstBincutCsharp = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Bincut_Csharp");
            if (testInstBincutCsharp != null)
            {
                instSheet!.AddRows(new ReadInstanceSheet().ReadSheet(testInstBincutCsharp).Rows);
            }
            ExcelWorksheet? testInstVddbinningCs = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Vddbinning_CS");
            if (testInstVddbinningCs != null)
            {
                instSheet!.AddRows(new ReadInstanceSheet().ReadSheet(testInstVddbinningCs).Rows);
            }
            ExcelWorksheet? testInstBincutCsharpTsmc = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Bincut_CsharpTSMC");
            if (testInstBincutCsharpTsmc != null)
            {
                instSheet!.AddRows(new ReadInstanceSheet().ReadSheet(testInstBincutCsharpTsmc).Rows);
            }

            ExcelWorksheet? postTestInstSheet = isCsharp
                ? (
                    IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Post_BinCut_CS")
                    ?? IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Vddbinning_PostBV")
                    ?? IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Vddbinning_OutsideBV")
                )
                : (
                    IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Vddbinning_OutsideBV")
                    ?? IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "TestInst_Vddbinning_PostBV")
                );

            if (postTestInstSheet != null)
            {
                List<InstanceRow> instRows = new ReadInstanceSheet().GetSheetInstanceRows(postTestInstSheet);
                instSheet?.AddRows(instRows);

                richTextBoxAppend("Reading Post TestInst_Vddbinning ...", Color.Blue);
            }
            BinCutData.TestInstanceSheet = instSheet;
        }

        internal static void ReadTestProgram(bool isCsharp, Action<string, Color> richTextBoxAppend, string testProgramFilePath, string tempFolder)
        {
            richTextBoxAppend("Reading test program ...", Color.Blue);
            if (string.IsNullOrEmpty(testProgramFilePath))
            {
                throw new Exception("Please assign test program for datalog checking !!!");
            }
            var exportMain = new IgxlManager();
            richTextBoxAppend("Exporting test program to txt files ...", Color.Blue);
            IgxlManager.ExportWorkBook(testProgramFilePath, tempFolder);

            List<BinTableRow> binTableRows = LoadBinTableRows(testProgramFilePath, richTextBoxAppend);
            CheckBinCutFlags(binTableRows, richTextBoxAppend);

            ExcelWorksheet? pinMapSheet = IgxlSheetReaderHelpers.GetExcelWorksheet(testProgramFilePath, "PinMap");
            if (pinMapSheet != null)
            {
                richTextBoxAppend("Reading PinMap ...", Color.Blue);
                BinCutData.PinMapData = new ReadPinMapSheet().ReadSheet(pinMapSheet);
            }

            List<IIgxlSheet> jobList = IgxlSheetReaderHelpers.GetIgxlSheetsByIgxlFile(testProgramFilePath, EnumSheetType.DTJobListSheet);
            if (jobList.Count != 0)
            {
                richTextBoxAppend("Reading JobList ...", Color.Blue);
                var jobListSheet = jobList.FirstOrDefault() as JobListSheet;
                BinCutData.JoblistSheet = jobListSheet;
            }

            ReadInstanceSheet(testProgramFilePath, isCsharp, richTextBoxAppend);
        }
    }
}
