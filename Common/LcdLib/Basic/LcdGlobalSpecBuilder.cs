using System.Collections.Generic;

using Automation.GenerateIgxl.Basic.Business.GenGlobalDc.Business;
using Automation.GenerateIgxl.Basic.Business.GenGlobalSpec;
using Automation.InputManager.Data;
using Automation.Singleton;
using Automation.Static;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.DataStruct;

namespace LcdLib.Basic
{
    internal static class LcdGlobalSpecBuilder
    {
        internal static GlobalSpecSheet Generate(BasicInputData basicInputData, MultiTestSettingSheetsSingleton multiTestSettingSheetsSingleton, IoLevelsSheet ioLevelsSheet)
        {
            var glbSpecSheet = new GlobalSpecSheet("Global Specs");
            GlobalSpecGenerator globalGenerator = new GlobalSpecGenerator(basicInputData.PinMapSheet, basicInputData.IoContiSheet);
            List<GlobalSpec> powerGlbs = globalGenerator.GetPowerGlobalSpecs(TestPlanStatic.PowerInfoSheet, multiTestSettingSheetsSingleton, basicInputData.IfoldPowerTableSheet);
            glbSpecSheet.AddRows(powerGlbs);
            glbSpecSheet.AddRow(new GlobalSpec("IO_Pins_GLB_Plus", "=1.1"));
            glbSpecSheet.AddRow(new GlobalSpec("IO_Pins_GLB_Minus", "=0.9"));
            glbSpecSheet.AddRows(ioLevelsSheet.GenGlbSymbol());

            GlbSymbolGenerator glbSymbolPlus = new GlbSymbolGenerator();
            glbSymbolPlus.DefaultGlbSymbols(glbSpecSheet);
            glbSymbolPlus.PlusGlbSymbols(glbSpecSheet);

            return glbSpecSheet;
        }
    }
}
