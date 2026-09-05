using System.Collections.Immutable;

using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo.FunctionAssigner.Universal;
using Automation.InputManager.Data;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetUniversalValue : SetVifValue
    {
        public SetUniversalValue(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet) : base(hardIpInputData, hardIpSheet)
        {
            ReservedMiscInfoKeys =
            [
                "digCapDataCustomString",
                "SelSramBlock",
                "BypassAutorange",
                "RAK",
                "ConvertDAC",
                "ConvertDACParameter",
                "SignedBinary",
                "RAK",
                "BypassAutorange",
                "DDREmulate",
                "UnSignedGray",
                "SignedGray",
                "TwosComplement",
                "SweepSrc",
                "SweepSrcOneToOne",
                "SweepVoltage",
                "DcvsVoltageClamp",
                "DigCapMsbFirst",
                "DigSrcMsbFirst",
                "RunByModule"
            ];
        }

        private static readonly ImmutableList<IFunctionAssigner> _functionAssigners =
        [
            new PatSetAssigner(),
            new CpuFlagAssigner(),
            new RegAssignAssigner(),
            new DigiCapAssigner(),
            new MeasCAssigner(),
            new MeasPinsAssigner(),
            new SpecialCalcAssigner(),
            new InterposeAssigner(),
            new MiscAssigner(),
        ];

        public override void SetArgsListValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            FunctionAssignContext ctx = new(pattern, voltage, HardIpInputData);
            foreach (IFunctionAssigner functionAssigner in _functionAssigners)
            {
                functionAssigner.Assign(function, ctx);
            }
        }
    }
}
