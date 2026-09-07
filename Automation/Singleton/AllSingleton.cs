using System.Collections.Generic;

using Automation.GenerateIgxl.PreAction.GenDSSCSetup;
using Automation.Static;

using TestPlanLib.Singleton;

namespace Automation.Singleton
{
    public static class AllSingleton
    {
        public static void Initialize()
        {
            AcTSetCategoryMapSingleton.Initialize(LocalSpecs.TimeSetFolder, LocalSpecs.Options.IsIgnorePatternCheck);
            PatternListSingleton.Initialize();
            CharSetupSingleton.Initialize();
            List<string> jobs = LocalSpecs.AllJobs;
            MultiTestSettingSheetsSingleton.Initialize(EpWorkbook.TestPlanWorkbook, LocalSpecs.CurrentProject, jobs, LocalSpecs.Options.Device);
            BinNumberSingletonLegacy.Initialize(LocalSpecs.ConfigFiles.BinNumberConfig, LocalSpecs.CurrentProject);
            NwireSingleton.Initialize();
            SelSramPatternSingleton.Initialize();
            SelSramWriterSingleton.Initialize();
            BinNumberSingleton.Instance.Initialize(EpWorkbook.TestPlanWorkbook);
            CharSetupSheetSingleton.Initialize();
            IdsRepairSingleton.Initialize();
            DsscSetupMain.Initialize();
            ExceptionListSingleton.Instance().Initialize();
            PerformanceModeSingleton.Initialize();
        }
    }
}
