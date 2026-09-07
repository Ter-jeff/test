using System;
using System.Collections.Generic;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.ErrorReport;
using CommonLib.ErrorReport.ErrorCodes;
using CommonLib.Extension;
using CommonLib.Utility;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Efuse.Input;

namespace Automation.GenerateIgxl.HardIp.HardIpPreCheck
{
    public class HardIpEfuseFieldChecker
    {
        private readonly Dictionary<string, HardIpSheet> _planDic = new Dictionary<string, HardIpSheet>();
        public HardIpEfuseFieldChecker(Dictionary<string, HardIpSheet> planDic)
        {
            _planDic = planDic.Where(x => x.Key.ToUpper().StartsWith("HARDIP_")).ToDictionary(x => x.Key, y => y.Value);
        }

        public void Check()
        {
            //Skip IDS and RTOS filed check
            if (!_planDic.Keys.Any())
            {
                return;
            }

            const string eFuseHardIpTable = "eFuse_HardIP_Table";
            ExcelWorksheets testPlanSheets = EpWorkbook.TestPlanWorkbook.Worksheets;
            ExcelWorksheet efuseBitDefSheet = GetEfuseBitDefSheet(testPlanSheets, SheetConst.Type5BitDefTable);
            ExcelWorksheet eFuseHardIpTableSheet = GetEfuseBitDefSheet(testPlanSheets, eFuseHardIpTable);
            List<string> eFuseHardIpCateNameList = CollectCateNameFromEfuseHardIp(_planDic);
            string labelCheckingItemName = "";
            try
            {
                //Read EFUSE_BitDef_Table

                List<BitDefTable> efuseBitDefTables = null;
                if (efuseBitDefSheet != null)
                {
                    efuseBitDefTables = TestPlanStatic.BitDefTables;
                    //load efuse bit def table for HIPEfuseRead/HIPEfuseWrite check
                    labelCheckingItemName = "EfuseBitTableCheker";
                    //GetEfuseBwDataForCheck(efuseBitDefTables);
                }

                EFuseHardIpSheet eFuseHardIpSheet = null;
                if (eFuseHardIpTableSheet != null)
                {
                    eFuseHardIpSheet = new EFuseHardIpReader().ReadSheet(eFuseHardIpTableSheet);
                }
                if (efuseBitDefTables != null && eFuseHardIpSheet != null)
                {
                    //check efuse hard ip table with efuse bit def
                    labelCheckingItemName = "HIPEfuseFieldChecker";
                    CheckEfuseHardIpAndEfuseBitDefTable(efuseBitDefTables, eFuseHardIpSheet, eFuseHardIpCateNameList);
                    CheckEfuseRealValue(efuseBitDefTables, eFuseHardIpSheet);
                }
                else
                {
                    if (efuseBitDefTables == null)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_01, SheetConst.Type5BitDefTable, 0, 0);
                    }
                    if (eFuseHardIpTableSheet == null)
                    {
                        ErrorReportManager.AddError(HardIpErrorType.E_HIPeFuseInput_02, eFuseHardIpTable, 0, 0);
                    }
                }
            }
            catch (Exception e)
            {
                string outString = "Check failed at " + labelCheckingItemName + "-----" + e.StackTrace;
                Response.Report(outString, EnumMessageLevel.Error, 100);
            }
        }

        internal List<string> CollectCateNameFromEfuseHardIp(Dictionary<string, HardIpSheet> planDic)
        {
            var mCateNameList = new List<string>();
            foreach (string key in planDic.Keys)
            {
                foreach (HardIpPattern pattern in planDic[key].Rows)
                {
                    if (pattern.MiscInfo.ContainsIgnoreCase("HIP_EFUSE_READ") || pattern.MiscInfo.ContainsIgnoreCase("HIP_EFUSE_WRITE") ||
                        pattern.MiscInfo.ContainsIgnoreCase("HardIPFuseRead") || pattern.MiscInfo.ContainsIgnoreCase("HardIPFuseWrite"))
                    {
                        //get the catename from the testplan
                        foreach (string argrument in pattern.MiscInfo.Split(';'))
                        {
                            if (string.IsNullOrEmpty(argrument) || argrument.Split(':').Length < 1)
                            {
                                continue;
                            }

                            if (!argrument.Contains(':'))
                            {
                                continue;
                            }

                            string funcKey = argrument.Split(':')[0];
                            string value = argrument.Split(':')[1];
                            if (funcKey.Equals("m_catename", StringComparison.CurrentCultureIgnoreCase))
                            {
                                mCateNameList.AddRange(value.Split('+'));
                            }
                        }
                    }
                }
            }

            return mCateNameList;

        }

        internal void CheckEfuseRealValue(List<BitDefTable> efuseBitDefTables, EFuseHardIpSheet eFuseHardIpSheet)
        {
            //Check BDF real code is match to efuse_hardIP_table 
            foreach (BitDefRow efuseRow in efuseBitDefTables.SelectMany(p => p.Rows).Where(x => x.RowData[efuseBitDefTables[0].DefaultOrRealIdx].Equals("real", StringComparison.CurrentCultureIgnoreCase)))
            {
                if (eFuseHardIpSheet.Rows.FirstOrDefault(x => x.Efusedefinition.Equals(efuseRow.RowData[efuseBitDefTables[0].BankEfuseBitDefIdx])) == null)
                {
                    string outString = $"{efuseRow.RowData[efuseBitDefTables[0].BankEfuseBitDefIdx]} is real value in Efuse_Bit_Def but can not find in eFuse_HardIP_Table.";
                    ErrorReportManager.AddError(
                        HardIpErrorType.W_MissingRealFieldNameInEfuseHardIP_01,
                        efuseRow.SheetName,
                        efuseRow.RowNum,
                        efuseBitDefTables[0].BankEfuseBitDefIdx,
                        [efuseRow.RowData[efuseBitDefTables[0].BankEfuseBitDefIdx]]
                    );
                }
            }
        }

        internal void CheckEfuseHardIpAndEfuseBitDefTable(List<BitDefTable> efuseBitDefTables, EFuseHardIpSheet eFuseHardIpSheet, List<string> eFuseHardIpCateNameList)
        {
            foreach (EFuseHardIpRow hardipRow in eFuseHardIpSheet.Rows)
            {
                // check the prewrite cate name is defined in the hardip testplan
                bool preWriteExistInTestplan = eFuseHardIpCateNameList.FirstOrDefault(x => x.Equals(hardipRow.Efusedefinition, StringComparison.CurrentCultureIgnoreCase)) != null;
                if (!preWriteExistInTestplan)
                {
                    string outString = $"The field {hardipRow.Efusedefinition} is not defined in the prewrite HardIP testplan.";
                    ErrorReportManager.AddError(
                        HardIpErrorType.E_MissingInPreWrite_01,
                        eFuseHardIpSheet.SheetName,
                        hardipRow.RowNum,
                        eFuseHardIpSheet.HeaderIndex["Width"],
                        [hardipRow.Efusedefinition]
                    );
                }
                // check bank
                List<BitDefTable> bankItems = efuseBitDefTables.FindAll(x => x.BlockName.Replace("_", "").Equals(hardipRow.Bank.Replace("_", ""), StringComparison.CurrentCultureIgnoreCase));
                if (bankItems != null)
                {
                    BitDefRow target = bankItems.SelectMany(x => x.Rows).FirstOrDefault(x => x.RowData[efuseBitDefTables[0].BankEfuseBitDefIdx].Equals(hardipRow.Efusedefinition, StringComparison.CurrentCultureIgnoreCase));
                    if (target != null)
                    {
                        try
                        {
                            int.TryParse(target.RowData[efuseBitDefTables[0].BitWidthIdx], out int widthEfuse);
                            int.TryParse(hardipRow.Width, out int widthHipEfuse);
                            if (widthHipEfuse != widthEfuse)
                            {
                                string outString = $"The bit width of the {hardipRow.Efusedefinition} is not match to the Efuse_Bit_Def definition.";
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_MismatchFieldWidth_01,
                                    eFuseHardIpSheet.SheetName,
                                    hardipRow.RowNum,
                                    eFuseHardIpSheet.HeaderIndex["Width"],
                                    [hardipRow.Efusedefinition]
                                );
                            }

                            if (!target.RowData[efuseBitDefTables[0].ProgrammingStageIdx].Equals(hardipRow.Job, StringComparison.CurrentCultureIgnoreCase))
                            {
                                string outString = $"The job of the {hardipRow.Efusedefinition} is not match to the Efuse_Bit_Def definition.";
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_MismatchFieldJob_01,
                                    eFuseHardIpSheet.SheetName,
                                    hardipRow.RowNum,
                                    eFuseHardIpSheet.HeaderIndex["Job"],
                                    [hardipRow.Efusedefinition]
                                );
                            }
                            if (!target.RowData[efuseBitDefTables[0].DefaultOrRealIdx].Equals("real", StringComparison.CurrentCultureIgnoreCase))
                            {
                                string outString = $"The field of the {hardipRow.Efusedefinition} is not real, please check with the Efuse_Bit_Def definition.";
                                ErrorReportManager.AddError(
                                    HardIpErrorType.E_MismatchRealValue_01,
                                    target.SheetName,
                                    target.RowNum,
                                    efuseBitDefTables[0].DefaultOrRealIdx,
                                    [hardipRow.Efusedefinition]
                                );
                            }
                        }
                        catch (Exception e)
                        {
                            string outString = "----- Exception when HardIP Efuse Field checker -----" + e.StackTrace;
                            Response.Report(outString, EnumMessageLevel.Error, 100);
                        }
                    }
                    else
                    {
                        ErrorReportManager.AddError(
                            HardIpErrorType.E_MissingFieldName_01,
                            eFuseHardIpSheet.SheetName,
                            hardipRow.RowNum,
                            eFuseHardIpSheet.HeaderIndex["eFuse Definition"],
                            [hardipRow.Efusedefinition]
                        );
                    }
                }
            }
        }

        internal static ExcelWorksheet GetEfuseBitDefSheet(ExcelWorksheets testPlanSheets, string sheetName)
        {
            foreach (ExcelWorksheet sheet in testPlanSheets)
            {
                if (CellDiff.IsLiked(sheet.Name, sheetName))
                {
                    return sheet;
                }
            }
            return null;
        }
    }
}
