using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.PreAction.GenDSSCSetup;
using Automation.Singleton;
using Automation.Static;

using CommonLib.Enums;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

using ProjectConfigLib.ProjectConfig;

using TestPlanLib.DataStruct;
using TestPlanLib.Singleton;
using TestPlanLib.Static;

namespace Automation.Test.UT
{
    public class FunctionTestBase : TestBase
    {
        [ClassInitialize(InheritanceBehavior.BeforeEachDerivedClass)]
        public static void ClassInitialize(TestContext testContext)
        {
            MockService.Mock();
            AllStatic.Clear();
            NeededSheets.InitSheetName(EnumDevice.AP, Directory.GetCurrentDirectory(), "");
            LocalSpecs.IsUnitTest = true;
            LocalSpecs.Options.Device = EnumDevice.AP;
            BorneoInitialize();
        }

        public static void BorneoInitialize()
        {
            LocalSpecs.CurrentProject = "borneo";
            ProjectConfigSingleton.Instance().InitializeProjectConfigSetting();
            ProjectConfigSingleton.Instance().LoadProjectConfig(LocalSpecs.ProjectIniFileName);
            LocalSpecs.Options = new Options(ProjectConfigSingleton.Instance());
            LocalSpecs.EquationVoltagesFileName = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "Borneo_A0_EQN_V04A_X_X#1.xls");

            LocalSpecs.BasLibraryFolder = Path.Combine(KPath, "central_library_vb");
            LocalSpecs.CsLibraryFolder = Path.Combine(KPath, "Lib");
            LocalSpecs.PatternFolder = Path.Combine(KPath, "borneo", "patx");
            LocalSpecs.TestPlanFileName = Path.Combine(InputPath, "borneo_documents", "borneo_A0_TestPlan.xlsx");
            MultiTestSettingSheetsSingleton.Initialize(EpWorkbook.TestPlanWorkbook, LocalSpecs.CurrentProject, LocalSpecs.AllJobs, EnumDevice.AP);
            NwireSingleton.Initialize();
            CharSetupSingleton.Initialize();
            PatternListSingleton.Initialize();
            SelSramPatternSingleton.Initialize();
            SelSramWriterSingleton.Initialize();
            CharSetupSheetSingleton.Initialize();
            IdsRepairSingleton.Initialize();
            DsscSetupMain.Initialize();

            LocalSpecs.ScghFileName = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "borneo_A0_SCGH_X_X_X#208.xlsx");
            NeededSheets.InitSheetName(EnumDevice.AP, LocalSpecs.SettingFolder, LocalSpecs.CurrentProject);

            LocalSpecs.TimeSetFolder = Path.Combine(KPath, "borneo", "TimeSet");
            LocalSpecs.PatternListCsvFileName = Path.Combine(InputPath, "borneo_documents", "A0_V04A", "borneo_A0_pattern_dashboard_20251002_1524_15.csv");
            AcTSetCategoryMapSingleton.Initialize(LocalSpecs.TimeSetFolder, LocalSpecs.Options.IsIgnorePatternCheck);

            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            BinNumberSingletonLegacy.Initialize(LocalSpecs.ConfigFiles.BinNumberConfig, LocalSpecs.CurrentProject);
        }

        protected static (IoInfoSheet printSheet, IoInfoSheet printConcurrentSheet) GenIoInfoSheets()
        {
            Dictionary<string, List<IoInfoRow>> blockInfo = new Dictionary<string, List<IoInfoRow>>
            {
                ["HardIP"] = [new IoInfoRow { PinGrpName = "HardIP1", Nv = "A" }],
                ["Scan"] = [new IoInfoRow { PinGrpName = "Scan1", Nv = "A", Vih = "1", Vil = "0.5", Iol = "0", Ioh = "0.0001", Vcl = "-1.2", Vch = "1.2" }],
                ["DcTest_Continuity"] = [new IoInfoRow { PinGrpName = "Conti1", Nv = "A" }],
                ["Levels_WalkingZ_Pos"] = [new IoInfoRow { PinGrpName = "WalkingZ", Nv = "A", Iol = "0", Ioh = "0.0001", Vcl = "-1.2", Vch = "1.2" }],
                ["Levels_WalkingZ_Pos_85C"] = [new IoInfoRow { PinGrpName = "WalkingZ", Nv = "A", Iol = "0", Ioh = "0.0001", Vcl = "-1.3", Vch = "1.3" }],
                ["Levels_WalkingZ_Neg_85C"] = [new IoInfoRow { PinGrpName = "WalkingZ", Nv = "A", Iol = "0", Ioh = "-0.0001", Vcl = "-1.3", Vch = "1.3" }],
                ["Common"] = [new IoInfoRow { PinGrpName = "IO", Nv = "A" }],
                ["Levels_BinCut:Scan"] = [new IoInfoRow { PinGrpName = "Scan1", Nv = "A", Vil = "0.3", Iol = "0", Vcl = "-1.2", Vch = "1.2" }],
            };

            var printSheet = new IoInfoSheet("IoInfo", blockInfo);

            Dictionary<string, List<IoInfoRow>> concurrentBlockInfo = new Dictionary<string, List<IoInfoRow>>
            {
                ["HardIP"] = [new() { PinGrpName = "HardIP2", Nv = "A" }],
                ["Scan"] = [new() { PinGrpName = "Scan2", Nv = "A" }],
                ["DcTest_Continuity"] = [new() { PinGrpName = "Conti2", Nv = "A" }],
                ["Common"] = [new() { PinGrpName = "IO", Nv = "A" }]
            };
            var printConcurrentSheet = new IoInfoSheet("IOInfo_Concurrent", concurrentBlockInfo);
            return (printSheet, printConcurrentSheet);
        }
    }
}
