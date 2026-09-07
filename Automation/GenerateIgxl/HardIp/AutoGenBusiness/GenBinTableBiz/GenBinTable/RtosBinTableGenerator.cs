using System.Collections.Generic;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTableRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTableBiz.GenBinTable
{
    public class RtosBinTableGenerator : BlockBinTableGeneratorBase
    {
        public RtosBinTableGenerator(HardIpInputData hardIpInputData, BinTableSheet hardIpBinTableSheet, string sheetName, List<HardIpPattern> patternList, List<string> duplicateParameter, List<string> errorBinNums)
            : base(hardIpInputData, hardIpBinTableSheet, patternList, duplicateParameter)
        {
            BinTableRowGenerator = new RtosBinTableRowGenerator(sheetName, errorBinNums);
        }
        public override void GenerateBinTableRows(bool mergedHln = false)
        {
            var binTableList = new List<BinTableRow>();
            foreach (HardIpPattern pattern in PatternList)
            {
                //Bypass generate bin table for run sc
                //if (pattern.FunctionName.Equals(FuncNameConst.CSharpFuncNameRtosRunScenario, StringComparison.OrdinalIgnoreCase))
                //{
                //    continue;
                //}
                if (HardIpConstData.RegOpcodeInPatt.IsMatch(pattern.Pattern.RealPatternName) && string.IsNullOrEmpty(pattern.Failflag))
                {
                    continue;
                }

                string voltage = GetVoltageForBinTable(pattern);
                binTableList.Add(BinTableRowGenerator.GenBinTableRow(pattern, voltage));
            }
            if (!binTableList.Exists(p => DuplicateParameter.Contains(p.Name)))
            {
                foreach (BinTableRow item in binTableList)
                {
                    if (item != null)
                    {
                        HardIpBinTableSheet.AddRow(item);
                        DuplicateParameter.Add(item.Name);
                    }
                }
            }
        }
    }
}
