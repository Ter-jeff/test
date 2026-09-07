using System.Collections.Generic;

using Automation.InputManager.Data;

using IgxlLib.IgxlSheets;

namespace RfLib.Dvdc
{
    internal static class DvdcSheetsGenerator
    {
        internal static (
            List<TimeSetBasicSheet> TimeSetSheets,
            PatSetSheet PatSetSheet,
            WaveDefinitionSheet WaveDefinitionSheet,
            MixedSignalSheet MixedSignalSheet,
            List<InstanceSheet> InstSheets,
            List<SubFlowSheet> FlowSheets,
            BinTableSheet BinTableSheet,
            CharSheet CharSheet,
            SubFlowSheet InitFlow) Generate(HardIpInputData hardIpInputData)
        {
            DvdcPreProcessor.PreProcess(hardIpInputData);

            (List<TimeSetBasicSheet> timeSetSheets, PatSetSheet patSetSheet, WaveDefinitionSheet waveDefinitionSheet, MixedSignalSheet mixedSignalSheet, List<InstanceSheet> instSheets, List<SubFlowSheet> flowSheets, BinTableSheet binTableSheet, CharSheet charSheet) = DvdcCoreSheetsGenerator.Generate(hardIpInputData);

            SubFlowSheet initFlow = DvdcPostSheetsGenerator.Generate(hardIpInputData, instSheets, binTableSheet);

            return (timeSetSheets, patSetSheet, waveDefinitionSheet, mixedSignalSheet, instSheets, flowSheets, binTableSheet, charSheet, initFlow);
        }
    }
}
