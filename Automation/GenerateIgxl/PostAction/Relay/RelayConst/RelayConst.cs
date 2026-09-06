using Automation.Static;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.PostAction.Relay.RelayConst
{
    public class OutputConst
    {
        public const string RelayInstName = "TestInst_Common";
        public const string InstArgWaiTime = "0.003";
        public Function Function { get; private set; }

        public OutputConst(string relayOnPins, string relayOffPins, string waitTime = InstArgWaiTime)
        {
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName("ControlRelay", "hardip", true);
            if (function.IsFound && function.Type.ToUpper() == ".NET")
            {
                Function = function;
                function.SetParamValue("relayOnPins", relayOnPins);
                function.SetParamValue("relayOffPins", relayOffPins);
                function.SetParamValue("waitTime", waitTime);
            }
            else
            {
                function = TestProgram.VbtFunctionLib.GetFunctionByName("Relay_Control", "hardip");
                Function = function;
                function.SetParamValue("relay_on", relayOnPins);
                function.SetParamValue("relay_off", relayOffPins);
                function.SetParamValue("WaitTime", waitTime);
            }
        }
    }
}
