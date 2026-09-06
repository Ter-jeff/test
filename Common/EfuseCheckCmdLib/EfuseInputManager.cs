using System.Collections.Generic;
using System.IO;
using System.Linq;

using Automation.Const;
using Automation.GenerateIgxl.EFuse.Business;
using Automation.GenerateIgxl.EFuse.InputChecker;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;
using CommonLib.Extension;

using LogLib.Static;

using OfficeOpenXml;

using TestPlanLib.Basic;
using TestPlanLib.BinCut.BinCutInstance;
using TestPlanLib.Efuse.Input;

namespace EfuseCheckCmdLib
{
    public class EfuseInputManager
    {
        public MultiTestSettingSheetsSingleton? MultiTestSettingSheetsSingleton;

        public EfuseArraySizeSheet? EfuseArraySizeSheet;
        public EfuseReadSheet? EfuseReadSheet;

        public List<BitDefTable> EfuseBitDefTables { get; set; } = [];
        public List<BitDefBankRange> EfuseBitDefBankRange = [];
        public List<EfuseConfigMainSheet> EfuseConfigMainSheets = [];
        public int EfuseDatabaseRevision = -1;

        public List<EfusePatternRow> EfusePatternRows = [];
        public ScghData? EfuseScghSheet;
        public List<BinCutInstanceSheet> EfuseInstanceSheets = [];

        public void ReadAndPreCheck()
        {
            #region Read new testSetting
            //if (EpWorkbook.TestplanWorkbook.Worksheets[NeededSheets.PowerMerge] != null)
            //{
            //    Response.Report("Reading Test Setting ...", MessageLevel.General, 20);
            //    MultiTestSettingSheetsSingleton = MultiTestSettingSheetsSingleton.GetInstance();
            //}
            #endregion

            #region Read efuse sheet in test plan
            Dictionary<string, PatternData> patternDataVersion = AcTSetCategoryMapSingleton.Instance().PatternList;
            HashSet<string> patInK = AcTSetCategoryMapSingleton.Instance().PatternPathListInK;
            ReadTestPlanSheet(out bool onlyUdrP);
            #endregion

            #region Read pattern row
            if (TestPlanStatic.EfuseInstanceSheets != null && TestPlanStatic.EfuseInstanceSheets.Count != 0)
            {
                Dictionary<string, int> timeSets = AcTSetCategoryMapSingleton.Instance().DicTimeSetVersion;
                var instanceClassfication = new EfusePatternClassification(onlyUdrP, EfuseBitDefBankRange);
                #region Read from Efuse instance sheets
                Response.Report("Reading Efuse Patterns in Efuse instance sheets ...", EnumMessageLevel.General, 20);
                foreach (BinCutInstanceSheet efuseInstanceSheet in TestPlanStatic.EfuseInstanceSheets)
                {
                    new EfuseInstanceSheetChecker().WorkFlow(efuseInstanceSheet, patternDataVersion, timeSets);
                    if (efuseInstanceSheet != null)
                    {
                        EfuseInstanceSheets.Add(efuseInstanceSheet);
                        efuseInstanceSheet.AddToErrorReport();
                        EfusePatternRows.AddRange(instanceClassfication.ClassificationFromInstanceSheet(efuseInstanceSheet, patternDataVersion, patInK, EfuseBitDefTables));
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
                        EfuseScghSheet = new ScghData().LoadEfuseFromHardIpBistScghData(scghWorkbook);
                        var scghClassfication = new EfusePatternClassification(onlyUdrP, EfuseBitDefBankRange);
                        EfusePatternRows = scghClassfication.ClassificationFromScgh(EfuseScghSheet, patternDataVersion, patInK, EfuseBitDefTables);
                    }
                }
                #endregion
            }
            #endregion

            //other check
            if (EfuseConfigMainSheets.Count != 0 && EfuseBitDefTables.Count != 0 && EfuseBitDefTables.Exists(x => EFuseConst.GetBankName(x.BlockName) == BankType.Cfg))
            {
                new EfuseConfigureAllChecker().WorkFlow(EfuseConfigMainSheets, EfuseBitDefTables);
                foreach (EfuseConfigMainSheet efuseConfigMainSheet in EfuseConfigMainSheets)
                {
                    efuseConfigMainSheet.AddToErrorReport();
                }
                new EfuseConfigureBdfChecker().WorkFlow(EfuseConfigMainSheets.First(), EfuseBitDefTables, EfusePatternRows);
            }
        }

        public void ReadForCheckScript()
        {
            ReadTestPlanSheet(out bool onlyUdrP);

            if (EfuseConfigMainSheets.Count != 0 && EfuseBitDefTables.Count != 0 && EfuseBitDefTables.Exists(x => EFuseConst.GetBankName(x.BlockName) == BankType.Cfg))
            {
                new EfuseConfigureAllChecker().WorkFlow(EfuseConfigMainSheets, EfuseBitDefTables);
                foreach (EfuseConfigMainSheet efuseConfigMainSheet in EfuseConfigMainSheets)
                {
                    efuseConfigMainSheet.AddToErrorReport();
                }
                new EfuseConfigureBdfChecker().WorkFlow(EfuseConfigMainSheets.First(), EfuseBitDefTables);
            }
        }

        private void ReadTestPlanSheet(out bool onlyUdrP)
        {
            onlyUdrP = false;
            ExcelWorksheet efuseArraySizesheet = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_ARRAYS_SIZE"];
            if (efuseArraySizesheet != null)
            {
                Response.Report($"Reading {efuseArraySizesheet.Name} ...", EnumMessageLevel.General, 20);
                var efuseArraySizeReader = new EfuseArraySizeReader();
                EfuseArraySizeSheet = efuseArraySizeReader.ReadSheet(efuseArraySizesheet);
                new EfuseArraySizeChecker().WorkFlow(EfuseArraySizeSheet!);
                EfuseArraySizeSheet!.AddToErrorReport();
                onlyUdrP = CheckUdrBank(EfuseArraySizeSheet);

            }
            else
            {
                Response.Report("EFUSE_ARRAYS_SIZE missing ...", EnumMessageLevel.Error, 20);
            }

            ExcelWorksheet efuseBitDefTable = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_BitDef_Table"];
            ExcelWorksheet efuseDatabaseRevision = EpWorkbook.TestPlanWorkbook.Worksheets["EFUSE_DATABASE_REVISION"];
            if (efuseDatabaseRevision != null)
            {
                var efuseDatabaseRevisionReader = new EfuseDatabaseRevisionReader();
                EfuseDatabaseRevision = efuseDatabaseRevisionReader.ReadSheet(efuseDatabaseRevision);
            }

            if (efuseBitDefTable != null)
            {
                Response.Report($"Reading {efuseBitDefTable.Name} ...", EnumMessageLevel.General, 20);
                EfuseBitDefTables = TestPlanStatic.BitDefTables;
                new EfuseBitDefTableChecker().WorkFlow(EfuseBitDefTables);
                if (EfuseBitDefTables.Count != 0)
                {
                    EfuseBitDefBankRange = Automation.InputManager.EFuseInputManager.GetBitDefEachBankBitRange(EfuseBitDefTables);
                }
            }
            else
            {
                Response.Report("EFUSE_BitDef_Table missing ...", EnumMessageLevel.Error, 20);
            }

            foreach (ExcelWorksheet sheet in EpWorkbook.TestPlanWorkbook.Worksheets)
            {
                if (sheet.Name.StartsWithIgnoreCase("BKM_"))
                {
                    Response.Report($"Reading {sheet.Name} ...", EnumMessageLevel.General, 20);
                    //EfuseBkmInfoTable = new EfuseBkmInfoReader(sheet);
                }
                else if (sheet.Name.StartsWithIgnoreCase("EFUSE_CONFIG_MAIN_"))
                {
                    Response.Report($"Reading {sheet.Name} ...", EnumMessageLevel.General, 20);
                    var efuseConfigMainReader = new EfuseConfigMainReader();
                    EfuseConfigMainSheet? efuseConfigMainSheet = efuseConfigMainReader.ReadSheet(sheet);
                    new EfuseConfigureMainChecker().WorkFlow(efuseConfigMainSheet!);
                    efuseConfigMainSheet!.AddToErrorReport();
                    if (efuseConfigMainSheet != null)
                    {
                        EfuseConfigMainSheets.Add(efuseConfigMainSheet);
                    }
                }
            }

            ExcelWorksheet efuseRead = EpWorkbook.TestPlanWorkbook.Worksheets["EfuseRead"];
            if (efuseRead != null)
            {
                Response.Report($"Reading {efuseRead.Name} ...", EnumMessageLevel.General, 20);
                var efuseReader = new EfuseReadSheetReader();
                EfuseReadSheet = efuseReader.ReadSheet(efuseRead);
                EfuseReadSheet!.AddToErrorReport();
            }
            else
            {
                Response.Report("EfuseRead missing ...", EnumMessageLevel.Error, 20);
            }
        }

        private static bool CheckUdrBank(EfuseArraySizeSheet efuseArraySizeSheet)
        {
            return !efuseArraySizeSheet.Rows.Any(x => !string.IsNullOrEmpty(x.FuseArray) &&
                                                 (x.FuseArray.EndsWithIgnoreCase(BankType.UdrP0) ||
                                                  x.FuseArray.EndsWithIgnoreCase(BankType.UdrP1)));
        }
    }
}
