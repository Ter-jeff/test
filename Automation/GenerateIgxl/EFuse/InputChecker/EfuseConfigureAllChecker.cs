using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;

using LogLib.Static;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.EFuse.InputChecker
{
    public class EfuseConfigureAllChecker
    {
        public void WorkFlow(List<EfuseConfigMainSheet> sheets, List<BitDefTable> bdfTables)
        {
            int maxBits = GetCfgMaxBitsFromBdf(bdfTables);
            CheckEachRowInfo(sheets);
            CheckLastConditionMsbValue(sheets, maxBits);
        }

        private int GetCfgMaxBitsFromBdf(List<BitDefTable> bdfTables)
        {
            int maxBits = 0;
            BitDefTable configBdf = bdfTables.Find(x => EFuseConst.GetBankName(x.BlockName).Equals(BankType.Cfg));
            if (configBdf != null && configBdf.Rows.Any())
            {
                var configureRows = configBdf.Rows.Where(x => x.RowData[configBdf.AlgorithmIdx].Equals("cond", StringComparison.CurrentCultureIgnoreCase)).ToList();
                if (configureRows.Any())
                {
                    maxBits = configureRows.Max(x => Convert.ToInt32(x.RowData[configBdf.MsbBitIdx]));
                }
            }

            if (maxBits == 0)
            {
                string alarmStr = "Unable to get the correct maximum bit, please check BitDefTable";
                Response.Report(alarmStr, EnumMessageLevel.Error, 100);
            }
            return maxBits;
        }

        private void CheckEachRowInfo(List<EfuseConfigMainSheet> sheets)
        {
            for (int i = 0; i < sheets.Count; i++)
            {
                EfuseConfigMainSheet sheet1 = sheets[i];
                for (int j = i + 1; j < sheets.Count; j++)
                {
                    EfuseConfigMainSheet sheet2 = sheets[j];
                    List<EfuseConfigMainRow> rows1 = sheet1.Rows;
                    List<EfuseConfigMainRow> rows2 = sheet2.Rows;
                    if (rows1.Count != rows2.Count)
                    {
                        //string msg = $"The count of rows in sheet {sheet1.SheetName} ({rows1.Count}) is inconsistent with sheet {sheet2.SheetName} ({rows2.Count})";
                        sheet1.AddError(EFuseErrorType.E_MismatchRow_01, sheet1.SheetName, 0, 0, $"The count of rows in sheet {sheet1.SheetName} ({rows1.Count}) is inconsistent with sheet {sheet2.SheetName} ({rows2.Count})", new string[] { sheet1.SheetName, rows1.Count.ToString(), $"{sheet2.SheetName}", $"{rows2.Count}" });
                        ErrorReportManager.AddError(EFuseErrorType.E_MismatchRow_01, sheet1.SheetName, 0, 0, [sheet1.SheetName, rows1.Count.ToString(), $"{sheet2.SheetName}", $"{rows2.Count}"]);
                    }
                }
            }
        }

        private void CheckLastConditionMsbValue(IEnumerable<EfuseConfigMainSheet> sheets, int maxBit)
        {
            foreach (EfuseConfigMainSheet sheet in sheets)
            {
                if (!sheet.Rows.Exists(x => x.Description.StartsWith("CFG_condition_", StringComparison.CurrentCultureIgnoreCase) && x.Msb >= maxBit))
                {
                    if (
                        sheet.Rows.Any(y => y.Description.StartsWith("CFG_condition_", StringComparison.CurrentCultureIgnoreCase)))
                    {
                        int msbMax =
                            sheet.Rows.Where(y => y.Description.StartsWith("CFG_condition_", StringComparison.CurrentCultureIgnoreCase)).Max(x => x.Msb);
                        string msg = $"The last config condition MSB ({msbMax}) < Maximum Bit ({maxBit})";
                        ErrorReportManager.AddError(EFuseErrorType.E_InvalidMaximumBits_01, sheet.SheetName, 0, 0, [msbMax.ToString(), maxBit.ToString()]);
                    }
                }
            }
        }
    }
}
