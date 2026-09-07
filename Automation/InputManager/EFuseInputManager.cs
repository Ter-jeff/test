using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.EFuse.InputChecker;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace Automation.InputManager
{
    public class EFuseInputManager : InputManagerBase<EFuseInputData>
    {
        public EFuseInputManager(ExcelWorkbook excelWorkbook) : base(excelWorkbook)
        {
        }

        public override EFuseInputData Read()
        {
            var result = new EFuseInputData();

            #region Read efuse sheet in test plan
            Dictionary<string, PatternData> patternDataVersion = AcTSetCategoryMapSingleton.Instance().PatternList;
            HashSet<string> patInK = AcTSetCategoryMapSingleton.Instance().PatternPathListInK;
            bool onlyUdrP = false;
            ExcelWorksheet worksheet = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_ARRAYS_SIZE"];
            if (worksheet != null)
            {
                Response.Report($"Reading {worksheet.Name} ...", EnumMessageLevel.General, 20);
                var efuseArraySizeReader = new EfuseArraySizeReader();
                EfuseArraySizeSheet efuseArraySizeSheet = efuseArraySizeReader.ReadSheet(worksheet);
                new EfuseArraySizeChecker().WorkFlow(efuseArraySizeSheet);
                efuseArraySizeSheet.AddToErrorReport();
                onlyUdrP = CheckUdrBank(efuseArraySizeSheet);
                result.EfuseArraySizeSheet = efuseArraySizeSheet;
            }
            else
            {
                Response.Report("EFUSE_ARRAYS_SIZE missing ...", EnumMessageLevel.Error, 20);
            }

            ExcelWorksheet efuseDatabaseRevision = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_DATABASE_REVISION"];
            if (efuseDatabaseRevision != null)
            {
                var efuseDatabaseRevisionReader = new EfuseDatabaseRevisionReader();
                result.EfuseDatabaseRevision = efuseDatabaseRevisionReader.ReadSheet(efuseDatabaseRevision);
            }

            if (result.EfuseBitDefTables != null)
            {
                new EfuseBitDefTableChecker().WorkFlow(result.EfuseBitDefTables);
                if (result.EfuseBitDefTables.Any())
                {
                    result.EfuseBitDefBankRange = GetBitDefEachBankBitRange(result.EfuseBitDefTables);
                }
            }
            else
            {
                Response.Report("EFUSE_BitDef_Table missing ...", EnumMessageLevel.Error, 20);
            }

            foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWith("BKM_", StringComparison.CurrentCultureIgnoreCase))
                {
                    Response.Report($"Reading {sheet.Name} ...", EnumMessageLevel.General, 20);
                    result.EfuseBkmInfoTable = new EfuseBkmInfoReader(sheet);
                }
                else if (sheet.Name.StartsWith("EFUSE_CONFIG_MAIN_", StringComparison.CurrentCultureIgnoreCase))
                {
                    Response.Report($"Reading {sheet.Name} ...", EnumMessageLevel.General, 20);
                    var efuseConfigMainReader = new EfuseConfigMainReader();
                    EfuseConfigMainSheet efuseConfigMainSheet = efuseConfigMainReader.ReadSheet(sheet);
                    new EfuseConfigureMainChecker().WorkFlow(efuseConfigMainSheet);
                    efuseConfigMainSheet.AddToErrorReport();
                    result.EfuseConfigMainSheets.Add(efuseConfigMainSheet);
                }
            }

            ExcelWorksheet efuseRead = EpWorkbook.TestPlanWorkbook.Worksheets["EfuseRead"];
            if (efuseRead != null)
            {
                Response.Report($"Reading {efuseRead.Name} ...", EnumMessageLevel.General, 20);
                var efuseReader = new EfuseReadSheetReader();
                result.EfuseReadSheet = efuseReader.ReadSheet(efuseRead);

                result.EfuseReadSheet.AddToErrorReport();
            }
            else
            {
                Response.Report("EfuseRead missing ...", EnumMessageLevel.Error, 20);
            }

            #endregion

            #region Read pattern row
            if (TestPlanStatic.EfuseInstanceSheets != null && TestPlanStatic.EfuseInstanceSheets.Any())
            {
                Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
                var efusePatternClassification = new EfusePatternClassification(onlyUdrP, result.EfuseBitDefBankRange);
                #region Read from Efuse instance sheets
                Response.Report("Reading Efuse Patterns in Efuse instance sheets ...", EnumMessageLevel.General, 20);
                foreach (BinCutInstanceSheet efuseInstanceSheet in TestPlanStatic.EfuseInstanceSheets)
                {
                    new EfuseInstanceSheetChecker().WorkFlow(efuseInstanceSheet, patternDataVersion, timeSets);
                    if (efuseInstanceSheet != null)
                    {
                        efuseInstanceSheet.AddToErrorReport();
                        result.EfusePatternRows.AddRange(efusePatternClassification.ClassificationFromInstanceSheet(efuseInstanceSheet, patternDataVersion, patInK, result.EfuseBitDefTables));
                    }
                }
                #endregion
            }
            else
            {
                #region Read from SCGH
                if (LocalSpecs.ScghFileName != "N/A" && !string.IsNullOrEmpty(LocalSpecs.ScghFileName))
                {
                    if (File.Exists(LocalSpecs.ScghFileName))
                    {
                        Response.Report("Reading Efuse Patterns in SCGH ...", EnumMessageLevel.General, 20);
                        var excel = new ExcelPackage(new FileInfo(LocalSpecs.ScghFileName));
                        ExcelWorkbook scghWorkbook = excel.Workbook;
                        result.EfuseScghSheet = new ScghData().LoadEfuseFromHardIpBistScghData(scghWorkbook);
                        var efusePatternClassification = new EfusePatternClassification(onlyUdrP, result.EfuseBitDefBankRange);
                        result.EfusePatternRows = efusePatternClassification.ClassificationFromScgh(result.EfuseScghSheet, patternDataVersion, patInK, result.EfuseBitDefTables);
                    }
                }
                #endregion
            }
            #endregion

            return result;
        }

        public static List<BitDefBankRange> GetBitDefEachBankBitRange(IEnumerable<BitDefTable> bitDefTables)
        {
            var bitDefEachBankBitRange = new List<BitDefBankRange>();
            foreach (BitDefTable bitDefTable in bitDefTables)
            {
                string bankName = EFuseConst.GetBankName(bitDefTable.BlockName);
                if (bankName == BankType.Unknow)
                {
                    continue;
                }


                double firstLsb = Convert.ToDouble(bitDefTable.Rows.First().RowData[bitDefTable.LsbBitIdx]);
                double firstMsb = Convert.ToDouble(bitDefTable.Rows.First().RowData[bitDefTable.MsbBitIdx]);

                double lastLsb = Convert.ToDouble(bitDefTable.Rows.Last().RowData[bitDefTable.LsbBitIdx]);
                double lastMsb = Convert.ToDouble(bitDefTable.Rows.Last().RowData[bitDefTable.MsbBitIdx]);

                var bankRang = new BitDefBankRange
                {
                    BankName = bankName,
                    Lsb = firstMsb < firstLsb ? firstMsb : firstLsb,
                    Msb = lastLsb > lastMsb ? lastLsb : lastMsb
                };
                bankRang.Width = bankRang.Msb - bankRang.Lsb + 1;
                bitDefEachBankBitRange.Add(bankRang);
            }
            return bitDefEachBankBitRange;
        }

        private bool CheckUdrBank(EfuseArraySizeSheet arraySizeSheet)
        {
            return !arraySizeSheet.Rows.Any(x => !string.IsNullOrEmpty(x.FuseArray) &&
                                                 (x.FuseArray.EndsWith(BankType.UdrP0, StringComparison.CurrentCultureIgnoreCase) ||
                                                  x.FuseArray.EndsWith(BankType.UdrP1, StringComparison.CurrentCultureIgnoreCase)));
        }
    }
}
