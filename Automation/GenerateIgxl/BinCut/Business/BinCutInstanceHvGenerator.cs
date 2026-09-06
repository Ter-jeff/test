using System;
using System.Collections.Generic;

using Automation.GenerateIgxl.BinCut.Base;
using Automation.InputManager.Data;

using IgxlLib.IgxlBase;

namespace Automation.GenerateIgxl.BinCut.Business
{
    public class BinCutInstanceHvGenerator : BinCutInstanceLvGenerator
    {
        public BinCutInstanceHvGenerator(List<BinCutSourceItem> sourceRowList, BinCutInputData binCutInputManager, List<BinCutFinalInstanceRow> binCutFinalInstanceRows)
            : base(sourceRowList, binCutInputManager, binCutFinalInstanceRows)
        {
            Type = "HBV";
        }

        protected override List<FlowRow> GenBinTable(bool isPost, FlowRow row, BinCutRowForSort binCutRow, bool isCsharp = false)
        {
            var flowRows = new List<FlowRow>();
            if (!binCutRow.BinCutSourceRow.InstOrCallFlowByBms)
            {
                flowRows.Add(InstanceInterface.GetBinTableRow());
            }
            if (!isCsharp)
            {
                flowRows.Add(AddFailStopBinTable());
            }

            if (binCutRow.BinCutFinalInstanceRow.BinCutInstanceRow.BinOutStage.Equals("X", StringComparison.OrdinalIgnoreCase))
            {
                return flowRows;
            }

            if (isPost)
            {
                if (!ExistPostFlags.Exists(p => p.FlagName.Equals(row.FailAction, StringComparison.OrdinalIgnoreCase)))
                {
                    string name = InstanceInterface.GetBinTableName();
                    var binCutBinningItem = new BinCutBinningItem(name, "", binCutRow.BinCutSourceRow.TargetPerformanceMode, row.FailAction, binCutRow.BinCutSourceRow.ColumnName);
                    ExistPostFlags.Add(binCutBinningItem);
                }
            }
            else
            {
                if (!ExistHvccFlags.Exists(p => p.FlagName.Equals(row.FailAction, StringComparison.OrdinalIgnoreCase)) && !string.IsNullOrEmpty(row.FailAction))
                {
                    string name = InstanceInterface.GetBinTableName();
                    var binCutBinningItem = new BinCutBinningItem(name, binCutRow.BinCutSourceRow.GetDomainOfMode(),
                        binCutRow.BinCutSourceRow.TargetPerformanceMode, row.FailAction, binCutRow.BinCutSourceRow.ColumnName);
                    ExistHvccFlags.Add(binCutBinningItem);
                }
            }
            return flowRows;
        }
    }
}
