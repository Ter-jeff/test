using System.Collections.Generic;

using Automation.Singleton;

using CommonLib.Extension;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

using TestPlanLib.DataStruct;

namespace LcdLib.Basic
{
    internal static class LcdDcSpecSheetsBuilder
    {
        internal static List<DcSpecSheet> Generate(MultiTestSettingSheetsSingleton multiTestSettingSheetsSingleton, IoLevelsSheet ioLevelsSheet, GlobalSpecSheet globalSpecSheet)
        {
            DcSpecGeneratorLcd dcGenerator = new DcSpecGeneratorLcd(multiTestSettingSheetsSingleton, ioLevelsSheet, globalSpecSheet);
            Dictionary<string, List<DcSpec>> powerDcSpecs = dcGenerator.GetPowerDcSpecs();
            Dictionary<string, List<DcSpec>> ioDcSpecs = dcGenerator.GetIoDcSpecs();
            var multiDcSpecSheets = new List<DcSpecSheet>();

            foreach (KeyValuePair<string, List<DcSpec>> powerentry in powerDcSpecs)
            {
                string sheetName = "DC_Specs_" + powerentry.Key;

                List<string> selectorNameList = ["Min", "Typ", "Max"];
                TestSettingData? testSettingSheet = multiTestSettingSheetsSingleton.TestSettingSheetsList.Find(s => s.Job.EqualsIgnoreCase(powerentry.Key));
                List<string> categorys = testSettingSheet!.GetDcCategorys();
                DcSpecSheet powerDcSpecSheet = new DcSpecSheet(sheetName, categorys, selectorNameList);
                foreach (DcSpec dcSpec in powerentry.Value)
                {
                    powerDcSpecSheet.AddRow(dcSpec);
                }

                foreach (DcSpec dcSpec in ioDcSpecs[powerentry.Key])
                {
                    powerDcSpecSheet.AddRow(dcSpec);
                }

                multiDcSpecSheets.Add(powerDcSpecSheet);
            }

            return multiDcSpecSheets;
        }
    }
}
