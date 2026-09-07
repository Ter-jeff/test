using Automation.GenerateIgxl.HardIp.InputObject;
using Automation.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;
using Automation.GenerateIgxl.Wireless.DVDC.InstanceParameterSetting;
using Automation.InputManager.Data;
using Automation.Static;

using CommonLib.Enums;

using TestPlanLib.Static;
using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.HardIp.InstanceParameterSetting
{
    public class SetArgValueMain
    {
        private HardIpInputData HardIpInputData { get; }
        private HardIpSheet HardIpSheet { get; }

        public SetArgValueMain(HardIpInputData hardIpInputData, HardIpSheet hardIpSheet)
        {
            HardIpInputData = hardIpInputData;
            HardIpSheet = hardIpSheet;
        }

        public void SetArgsValue(HardIpPattern pattern, ref Function function, string voltage)
        {
            SetValueBase setValueFunction = ResolveSetValueFunction(function);

            setValueFunction.SetArgsListValue(pattern, ref function, voltage);
            setValueFunction.SetValueByParamMapping(function, pattern, voltage);
            // Need to check with lib if it needs to align the reg assign table header
            if (LocalSpecs.Options.Device == EnumDevice.LCD)
            {
                setValueFunction.CheckArgsExceedLimitationLcd(function, pattern);
            }
            else
            {
                if (!(function.FunctionName == VbtFunctionLibShared.RfFunc.ToLower()))
                {
                    setValueFunction.CheckArgsExceedLimitation(function, pattern);
                }
            }
            if (!(function.FunctionName == VbtFunctionLibShared.RfFunc.ToLower()))
            {
                setValueFunction.CheckInstArgument(function, pattern);
            }
        }

        private SetValueBase ResolveSetValueFunction(Function function)
        {
            string functionName = function.FunctionName.ToLower();

            SetValueBase typeDependent = ResolveTypeDependentFunction(functionName, function.Type);
            if (typeDependent != null)
            {
                return typeDependent;
            }

            if (functionName == VbtFunctionLibShared.FunctionalName.ToLower())
            {
                return new SetFunctionalValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.VirGpioTtrName.ToLower() || functionName == VbtFunctionLibShared.VirName.ToLower())
            {
                return new SetVirValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.Ids.ToLower())
            {
                return new SetIdsValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.IdsMathFunc.ToLower())
            {
                return new SetIdsMathFuncValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.LcdTrim.ToLower() || functionName == VbtFunctionLibShared.DvdcTrim.ToLower() || functionName == VbtFunctionLibShared.DvdcTrim3D.ToLower())
            {
                return new SetDvdcTrimValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.LcdMeas.ToLower())
            {
                return new SetLcdMeasValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.RfTrim.ToLower() || functionName == VbtFunctionLibShared.RfTrim2D.ToLower())
            {
                return new SetRfTrimValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.RfHtolFunc.ToLower())
            {
                return new SetRfHtolFuncValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.RfFunc.ToLower())
            {
                return new SetRfFuncValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == "HIP_eFuse_Read".ToLower() || functionName == "HardIPFuseRead".ToLower())
            {
                return new SetFuseReadValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == "HIP_eFuse_Write".ToLower() || functionName == "HardIPFuseWrite".ToLower())
            {
                return new SetFuseWriteValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == "RunScenario".ToLower())
            {
                return new SetRunScValue(HardIpInputData, HardIpSheet);
            }

            return new SetDefaultValue(HardIpInputData, HardIpSheet);
        }

        private SetValueBase ResolveTypeDependentFunction(string functionName, string functionType)
        {
            if (functionName == VbtFunctionLibShared.VifName.ToLower() && functionType == "VBT")
            {
                return new SetVifValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.VifName.ToLower() && functionType == ".NET")
            {
                return new SetUniversalValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.HardIpmtdTest.ToLower() && functionType == "VBT")
            {
                return new SetVifValue(HardIpInputData, HardIpSheet);
            }
            if (functionName == VbtFunctionLibShared.HardIpmtdTest.ToLower() && functionType == ".NET")
            {
                return new SetCsMtdfValue(HardIpInputData, HardIpSheet);
            }

            return null;
        }
    }
}
