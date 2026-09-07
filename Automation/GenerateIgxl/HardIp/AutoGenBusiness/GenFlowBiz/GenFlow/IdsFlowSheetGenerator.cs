using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlow
{
    public class IdsFlowSheetGenerator : HardIpFlowSheetGenerator
    {
        public IdsFlowSheetGenerator(HardIpInputData hardIpInputData, string sheetName, List<HardIpPattern> patternList) : base(hardIpInputData, sheetName, patternList)
        {
            FlowRowGenerator = new IdsFlowRowGenerator(hardIpInputData, SheetName);
        }

        protected override List<FlowRow> GenFlowBodyRows(bool shmooflag = false, bool vtShmooFlag = false)
        {

            var flowBodyRows = new List<FlowRow>();
            if (vtShmooFlag)
            {
                return flowBodyRows;
            }

            flowBodyRows.AddRange(GenFlowTestRowsByVoltage("NV", shmooflag, vtShmooFlag));
            flowBodyRows.AddRange(GenResetRelayRows());

            if (!shmooflag && !vtShmooFlag)
            {
                flowBodyRows.AddRange(GenFlowBinTableRows());
                flowBodyRows.AddRange(FlowRowGenerator.GenTtrFlagClearRow(flowBodyRows));
            }

            return flowBodyRows;
        }

        private List<FlowRow> GenFlowBinTableRows()
        {
            var flowBinTableRows = new List<FlowRow>();
            foreach (HardIpPattern pattern in ExtendedPatList)
            {
                FlowRowGenerator.Pat = pattern;
                if (pattern.IsIgnorePatBinOut() ||
                    Regex.IsMatch(pattern.Pattern.GetLastPayload(), "^cz_", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                FlowRow binRow = FlowRowGenerator.GenBinTableRow();
                flowBinTableRows.Add(binRow);
            }
            return flowBinTableRows;
        }
    }
}
