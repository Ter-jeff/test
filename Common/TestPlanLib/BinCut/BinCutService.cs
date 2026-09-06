using System.Collections.Generic;
using System.IO;

using CommonLib.Extension;

using OfficeOpenXml;

using TestPlanLib.BinCut.Binning;
using TestPlanLib.BinCut.NonIgxlSheet;
using TestPlanLib.Static;

namespace TestPlanLib.BinCut
{
    public static class BinCutService
    {
        public static string? GetEnumFromString(string input)
        {
            return input.ToUpper() switch
            {
                var s when s.Contains(nameof(JobStageEnum.CP1)) => nameof(JobStageEnum.CP1),
                var s when s.Contains(nameof(JobStageEnum.CP2)) => nameof(JobStageEnum.CP2),
                var s when s.Contains(nameof(JobStageEnum.FT1)) => nameof(JobStageEnum.FT1),
                var s when s.Contains(nameof(JobStageEnum.FT2)) => nameof(JobStageEnum.FT2),
                var s when s.Contains(nameof(JobStageEnum.FT3)) => nameof(JobStageEnum.FT3),
                _ => null,
            };
        }

        public static Dictionary<string, string> NonIgxlSheetProcess(Dictionary<string, string> modes, bool isCs, ExcelWorkbook excelWorkbook, string pOutFolder, bool isIds = false, string shadowStage = "")
        {
            double bcBaseVoltage = 0;
            double bcStepSize = 0;
            var result = new Dictionary<string, string>();

            #region Notes Sheet
            ExcelWorksheet notesWorksheet = excelWorkbook.Worksheets[NeededSheets.BinningNote];
            if (notesWorksheet != null)
            {
                var binningNotes = new BinCutNotes(notesWorksheet, pOutFolder);
                bcBaseVoltage = binningNotes.FindBaseVoltage();
                bcStepSize = binningNotes.FindStepSize();
            }
            #endregion

            #region Flow to *.txt
            ExcelWorksheet worksheet = excelWorkbook.Worksheets[NeededSheets.BcFlow] ?? excelWorkbook.Worksheets[NeededSheets.BcAte];
            if (worksheet != null)
            {
                var nonBinningRailSheet = new NonBinningRail(pOutFolder, worksheet);
                result.Add(nonBinningRailSheet.WorkFlow(isCs, shadowStage), NeededSheets.BcFlow);
            }
            #endregion

            #region Flow_PostBincut to *.txt
            int count = 0;
            foreach (ExcelWorksheet sheet in excelWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWithIgnoreCase("Flow_PostBincut"))
                {
                    count++;
                    var nonBinningRailPostSheet = new NonBinningRailPost(pOutFolder, sheet, count);
                    result.Add(nonBinningRailPostSheet.WorkFlow(isCs), sheet.Name);
                }
            }
            #endregion

            #region Binning to *.txt
            var binCutList = new List<string>();
            var binDomainSheetList = new List<string>();
            if (excelWorkbook.Worksheets[NeededSheets.Binning] != null)
            {
                binDomainSheetList.Add(NeededSheets.Binning);
                binCutList.Add("1");
            }

            if (excelWorkbook.Worksheets[NeededSheets.BinningBinX] != null)
            {
                binDomainSheetList.Add(NeededSheets.BinningBinX);
                binCutList.Add("2");
            }

            if (excelWorkbook.Worksheets[NeededSheets.BinningBinY] != null)
            {
                binDomainSheetList.Add(NeededSheets.BinningBinY);
                binCutList.Add("3");
            }

            for (int i = 0; i < binDomainSheetList.Count; i++)
            {
                worksheet = excelWorkbook.Worksheets[binDomainSheetList[i]];
                var vddBinningDefSheet = new VddBinningDef(pOutFolder, worksheet)
                {
                    BinCutList = string.Join(",", binCutList),
                    Index = i + 1,
                    BaseVoltage = bcBaseVoltage,
                    StepSize = bcStepSize
                };
                result.Add(vddBinningDefSheet.WorkFlow(isCs, shadowStage), binDomainSheetList[i]);
            }
            #endregion

            #region IDS distribution
            List<string> sVddBinList = [.. Directory.GetFiles(pOutFolder, !isCs ? "bincut_eqn_appA*.txt" : "Vdd_Binning_Def_appA*.txt")];
            if (sVddBinList.Count != 0)
            {
                BinningTable binningTable = BinningTableReader.Read(sVddBinList[0]);
                if (isIds)
                {
                    ExcelWorksheet sheet = excelWorkbook.Worksheets["EqnStart"];
                    var idsDistributionTxt = new IdsDistribution(pOutFolder, binningTable);
                    if (sheet != null)
                    {
                        var eqnStartSheetReader = new EqnStartSheetReader();
                        EqnStartSheet data = eqnStartSheetReader.ReadSheet(sheet)!;
                        result.Add(idsDistributionTxt.WorkFlow(data), "EqnStart");
                    }
                    else
                    {
                        result.Add(idsDistributionTxt.WorkFlow(), "IDS_Distribution");
                    }
                }
            }
            #endregion

            #region Power Binning to *.txt
            foreach (ExcelWorksheet sheet in excelWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWithIgnoreCase("PwrBin") || sheet.Name.StartsWithIgnoreCase("PwrScreen"))
                {
                    var nonBinningRailSheet = new BinCutNonIgxlBase(sheet, pOutFolder);
                    result.Add(nonBinningRailSheet.WorkFlow(isCs), sheet.Name);
                }
            }
            #endregion

            #region START_EQN to *.txt
            foreach (ExcelWorksheet sheet in excelWorkbook.Worksheets)
            {
                if (sheet.Name.EqualsIgnoreCase("START_EQN"))
                {
                    var startEqnSheet = new BinCutNonIgxlBase(sheet, pOutFolder);
                    result.Add(startEqnSheet.WorkFlow(isCs, shadowStage), sheet.Name);
                }
            }
            #endregion

            return result;
        }
    }
}
