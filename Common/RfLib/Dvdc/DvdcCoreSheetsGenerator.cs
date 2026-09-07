using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenCharBiz;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenInstanceBiz;
using Automation.InputManager.Data;
using Automation.Static;

using IgxlLib.IgxlSheets;

using LogLib.Static;

using RfLib.Dvdc.GenFlow;

namespace RfLib.Dvdc
{
    internal static class DvdcCoreSheetsGenerator
    {
        internal static (
            List<TimeSetBasicSheet> TimeSetSheets,
            PatSetSheet PatSetSheet,
            WaveDefinitionSheet WaveDefinitionSheet,
            MixedSignalSheet MixedSignalSheet,
            List<InstanceSheet> InstSheets,
            List<SubFlowSheet> FlowSheets,
            BinTableSheet BinTableSheet,
            CharSheet CharSheet) Generate(HardIpInputData hardIpInputData)
        {
            Response.Report("Generating TimeSet ...", percentage: 25);
            List<TimeSetBasicSheet> timeSetSheets = new TimeSetSheetGenerator().GenTimeSet(hardIpInputData.PlanDic);

            Response.Report("Generating Multiple Init PatSet ...", percentage: 30);
            PatSetSheet patSetSheet = new PatSetSheetGenerator(hardIpInputData).GenPatSet(hardIpInputData.PlanDic, ScghStatic.ScghData, "PatSets_RF");

            //Analog Setup, currently used to generate ADC autogen
            Response.Report("Generating AnalogSetup ...", percentage: 35);
            WaveDefinitionSheet waveDefinitionSheet = new HardIpWavedefGenerator().Generate(hardIpInputData.PlanDic);

            MixedSignalSheet mixedSignalSheet = new HardIpMixedSignalGenerator().Generate(hardIpInputData.PlanDic);

            Response.Report("Generating Instance ...", percentage: 45);
            List<InstanceSheet> instSheets = new InstanceGenerator(hardIpInputData).GenInst(hardIpInputData.PlanDic);

            Response.Report("Generating Testflow ...", percentage: 55);
            List<SubFlowSheet> flowSheets = new DvdcFlowGenerator(hardIpInputData).GenFlow(hardIpInputData.PlanDic);

            Response.Report("Generating Bin Table ...", percentage: 60);
            BinTableSheet binTableSheet = new BinTableSheetGenerator(hardIpInputData).GenBinTable(hardIpInputData.PlanDic);

            Response.Report("Generating Characterization ...", percentage: 70);
            CharSheet charSheet = new CharSheetGenerator().GenCharSheet(hardIpInputData.PlanDic);

            return (timeSetSheets, patSetSheet, waveDefinitionSheet, mixedSignalSheet, instSheets, flowSheets, binTableSheet, charSheet);
        }
    }
}
