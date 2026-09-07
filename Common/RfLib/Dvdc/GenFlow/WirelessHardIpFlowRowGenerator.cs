using System.Collections.Generic;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using Automation.GenerateIgxl.HardIp.AutoGenBusiness.GenFlowBiz.GenFlowRow;
using Automation.GenerateIgxl.HardIp.HardIpData.DataBase;
using Automation.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.InputManager.Data;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

namespace RfLib.Dvdc.GenFlow
{
    public partial class WirelessHardIpFlowRowGenerator(HardIpInputData hardIpInputData, string sheetName) : HardIpFlowRowGenerator(hardIpInputData, sheetName)
    {
        [GeneratedRegex("isFW", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex MyRegex();

        protected override void SetBasicInfoByPattern(HardIpPattern hardIpPattern)
        {
            BlockName = CommonGenerator.GetBlockNameFromSheetName(hardIpPattern.SheetName);
        }

        public override List<FlowRow> GenTestRows(bool isCz2Only = false)
        {
            var testRows = new List<FlowRow>();
            //testRow
            var testRow = new FlowRow
            {
                Job = CreateTestJob(),
                Part = CreateTestPart(),
                Opcode = CreateTestOpcode(),
                Env = CreateTestEnv(),
                Enable = isCz2Only ? "!ShmooOnly" : EnableWord,
                Parameter = CreateTestParameter(),
                FailAction = CreateTestFailAction()
            };
            testRows.Add(testRow);
            List<FlowRow> limitRows = SortFlowRows(GetLimitRows(testRow));
            testRows.AddRange(limitRows);
            return testRows;
        }

        public override List<FlowRow> GenShmooRows(string labelVoltage = "")
        {
            var testRows = new List<FlowRow>();
            if (!HasShmoo)
            {
                return testRows;
            }

            var testRow = new FlowRow
            {
                Enable = "!TestOnly",
                Opcode = CreateCharacterizeOpcode(),
                Parameter = CreateShmooParameter(),
                TName = CreateShmooTName()
            };

            testRows.Add(testRow);
            List<FlowRow> limitRows = SortFlowRows(GetLimitRows(testRow, forShmoo: true));
            testRows.AddRange(limitRows);
            return testRows;
        }

        protected override List<FlowRow> GetLimitRows(FlowRow flowRow, bool forShmoo = false)
        {
            string useLimitFailAction = CreateUseLimitFailAction(forShmoo);
            return GenUseLimitRows(flowRow.Parameter, useLimitFailAction);
        }

        protected override List<FlowRow> GenUseLimitRows(string parameter, string binFail)
        {
            var flowUseLimitRows = new List<FlowRow>();
            #region Generate Use-Limit rows
            List<MeasPin> useLimits = ActualLabelVoltage switch
            {
                HardIpConstData.LabelHv => Pat.UseLimitsH,
                HardIpConstData.LabelLv => Pat.UseLimitsL,
                HardIpConstData.LabelNv => Pat.UseLimitsN,
                _ => Pat.MeasPins,
            };
            foreach (MeasPin pin in useLimits)
            {
                if ((pin.MeasType.EqualsIgnoreCase(MeasType.MeasC) && pin.TestName.EqualsIgnoreCase("skip")) || pin.MeasType.EqualsIgnoreCase(MeasType.MeasN))
                {
                    continue;
                }

                var row = new FlowRow();
                string lowUnit;
                string lowScale;
                string highUnit;
                string highScale;
                row.Opcode = CreateUseLimitOpcode();
                row.Job = pin.Job;
                row.Parameter = parameter;
                if (!string.IsNullOrEmpty(pin.TestName))
                {
                    row.TName = pin.TestName;
                }

                if (MyRegex().IsMatch(Pat.MiscInfo))
                {
                    row.LoLim = DataConvertor.ConvertUseLimitFw(pin.LowLimit, out lowUnit, out lowScale);
                }
                else
                {
                    row.LoLim = DataConvertor.ConvertUseLimit(pin.LowLimit, out lowUnit, out lowScale);
                }

                if (MyRegex().IsMatch(Pat.MiscInfo))
                {
                    row.HiLim = DataConvertor.ConvertUseLimitFw(pin.HighLimit, out highUnit, out highScale);
                }
                else
                {
                    row.HiLim = DataConvertor.ConvertUseLimit(pin.HighLimit, out highUnit, out highScale);
                }

                row.Scale = lowScale.Length == 0 ? highScale : lowScale;
                row.Units = lowUnit.Length == 0 ? highUnit : lowUnit;
                row.FailAction = binFail;
                row.Comment = Pat.Pattern.GetLastPayload() + "_" + pin.MeasType + "_" + pin.PinName;
                flowUseLimitRows.Add(row);
            }
            #endregion
            return flowUseLimitRows;
        }
    }
}
