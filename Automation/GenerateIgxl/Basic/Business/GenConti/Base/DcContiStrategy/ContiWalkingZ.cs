using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.GenerateIgxl.Basic.Business.GenLevel.Business;
using Automation.GenerateIgxl.Basic.Business.GenTimeSet.Business;
using Automation.Reader;
using Automation.Static;
using Automation.Utility.Basic;
using Automation.Utility.Pattern;

using CommonLib.Enums;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiWalkingZ : ContiBase
    {
        private readonly Regex _regexDisconnectPair = new Regex(@"Disconnect[_]*Pair\s*=\s*(?<DisconnectPair>\w+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private readonly Regex _regexSetLowerPins = new Regex(@"Setlowpins\s*=\s*(?<SetLowPins>\w+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private string _disconnectPair = "";
        private string _setLowerPins = "";
        private readonly PatSetSheet _patSetAll;
        public ContiWalkingZ(DcTestContiRow dcTestContiRow, PatSetSheet patSetAll) : base(dcTestContiRow)
        {
            _patSetAll = patSetAll;
            SetCondition();
        }

        private bool HasDisconnectPair
        {
            get
            {
                return !string.IsNullOrEmpty(_disconnectPair);
            }
        }

        private void SetCondition()
        {
            Match m = _regexDisconnectPair.Match(DcTestContiRow.Condition);
            if (m.Success)
            {
                _disconnectPair = m.Groups["DisconnectPair"].Value;
            }

            m = _regexSetLowerPins.Match(DcTestContiRow.Condition);
            if (m.Success)
            {
                _setLowerPins = m.Groups["SetLowPins"].Value;
            }
        }

        public override List<FlowRow> GenerateFlowRows()
        {
            List<FlowRow> rowList = new List<FlowRow>();
            foreach (string job in DcTestContiRow.JobNameList)
            {
                string instanceName = DcTestContiRow.InstanceName + "_" + job;
                FlowRow newRow = CreateTestFlowRow(instanceName, "");
                string enable = $"{job}&&WalkingZ" + (string.IsNullOrEmpty(DcTestContiRow.EnableWord) ? "" : $"&&{DcTestContiRow.EnableWord}");
                newRow.Enable = enable;
                newRow.FailAction = CreateWalkingZFlagName(job);
                rowList.Add(newRow);

            }
            return rowList;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            List<InstanceRow> resultInstanceRows = new List<InstanceRow>();
            foreach (string job in DcTestContiRow.JobNameList)
            {
                InstanceRow row = new InstanceRow();
                string vbtName = DcContiConst.CSharpFuncNameWalkingZContinuity;
                row.TestName = DcTestContiRow.InstanceName + "_" + job;
                row.DcCategory = SetCategoryFromCondition();
                row.DcSelector = "Typ";
                row.TimeSets = TimeSetPlus.TsbContiPat;
                row.PinLevels = GetLevels();
                Function function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "conti", true);

                if (function.Type == ".NET")
                {
                    GenerateCSharpInstanceRow(ref function);
                }
                else
                {
                    vbtName = HasDisconnectPair ? DcContiConst.VbtContiWalkingZ : DcContiConst.VbtFuncNameFunctionalT;
                    function = TestProgram.VbtFunctionLib.GetFunctionByName(vbtName, "conti", true);
                    GenerateVbtInstanceRow(ref function);
                }
                SetArgumentFromCondition(ref function);

                row.VbtName = function.FullFunctionName;
                row.VbtType = function.Type;
                row.ArgList = function.Parameters;
                row.Args = function.ArgList;
                resultInstanceRows.Add(row);
                GenWalkingZPattern();
            }
            return resultInstanceRows;

        }

        private void GenerateCSharpInstanceRow(ref Function function)
        {
            function.SetParamValue("pattern", CreateWalkingZPatternName());
            function.SetParamValue("digitalPins", TestPinGroup);
            function.SetParamValue("differentialPairs", _disconnectPair);
            function.SetParamValue("setLowPins", _setLowerPins);
        }

        private void GenerateVbtInstanceRow(ref Function function)
        {
            function.SetParamValue("patset", CreateWalkingZPatternName());
            if (LocalSpecs.Options.Device != EnumDevice.LCD && function.FunctionName == DcContiConst.VbtFuncNameFunctionalT)
            {
                function.ArgList[24] = "1";
            }

            function.SetParamValue("digital_pins", _setLowerPins);
            function.SetParamValue("PN_Disconnect", _disconnectPair);
        }

        private string CreateWalkingZFlagName(string job)
        {
            return "F_WalkingZ_" + DcTestContiRow.InstanceName + "_" + job;
        }

        private void GenWalkingZPattern()
        {
            PinGroup testPins = TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(TestPinGroup);
            var allIOs = TestProgram.IgxlWorkBk.PinMapPair.Value.PinList.Where(a => a.PinType == PinMapConst.TypeIo).Select(a => a.PinName).ToList();
            List<string> diffPins = null;
            if (HasDisconnectPair)
            {
                PinGroup disconnectPair = TestProgram.IgxlWorkBk.PinMapPair.Value.GetGroup(_disconnectPair);
                if (disconnectPair != null)
                {
                    diffPins = disconnectPair.PinList.Select(x => x.PinName).ToList();
                }
            }
            else
            {
                diffPins = allIOs;
            }

            int repeatDummyCycles = 10;
            if (DcTestContiRow.ConditionDict.TryGetValue("repeatDummyCycles", out string value))
            {
                if (int.TryParse(value, out int tryGetValue) && tryGetValue > 0)
                {
                    repeatDummyCycles = tryGetValue;
                }
            }

            if (testPins != null && allIOs.Count > 0 && diffPins != null)
            {
                var testPinList = testPins.PinList.Select(x => x.PinName).ToList();
                Dictionary<string, string> diffPairs = DifferentialService.DifferentialPair(diffPins);
                string outputFolder = Path.Combine(LocalSpecs.TarFolder, "WalkingZPattern");
                var walkingZPat = new WalkingZPats(outputFolder, repeatDummyCycles);
                string patternName = CreateWalkingZPatternNameOnly(false);
                walkingZPat.WorkFlowFromContiSheet(allIOs, testPinList, patternName, diffPairs);
            }
        }

        protected string CreateWalkingZPatternName()
        {
            string patternName = CreateWalkingZPatternNameOnly(true);
            var patSet = new PatSet { PatSetName = patternName };
            patSet.AddRow(new PatSetRow { PatternSet = patternName, File = CreateWalkingZPatterPath(), Burst = "NO" });
            if (_patSetAll != null && !_patSetAll.IsExist(patternName))
            {
                _patSetAll.AddRow(patSet);
            }

            return patternName;
        }

        protected string CreateWalkingZPatterPath()
        {
            string patternName = CreateWalkingZPatternNameOnly(false);
            string dicPath = @".\PATTERN\WalkingZ\";
            return dicPath + patternName + ".PAT";
        }

        protected override string GetDefaultLevels()
        {
            string specs = "";
            foreach (string spec in new List<string> { "VOL", "VOH", "IOL", "IOH" })
            {
                if (DcTestContiRow.ConditionDict.TryGetValue(spec, out string value))
                {
                    specs += value;
                }
            }
            if (!string.IsNullOrEmpty(specs))
            {
                return specs.Contains("-") ? LevelInitial.WalkingZNegLevelSheetName : LevelInitial.WalkingZPosLevelSheetName;
            }
            return "";
        }
    }
}
