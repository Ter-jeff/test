using System;

using Automation;
using Automation.GenerateIgxl;
using Automation.Static;

using CommonLib.Enums;

using LcdLib.InputManager;
using LcdLib.InputManager.Data;
using LcdLib.OTP.Business;

using LogLib.Static;

using ParaData = Automation.PreCheck.AllParaData.ParaData;

namespace LcdLib.OTP
{
    public class OtpMain : WorkFlowBase<ParaData>
    {
        private OtpInputData? _otpInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                _otpInputData = new OtpInputManager(EpWorkbook.TestPlanWorkbook).Read();

                return true;
            }
            catch (Exception e)
            {
                Response.Report("Otp has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                var otpWorkFlowManager = new OtpWorkFlowManager(_otpInputData!);
                otpWorkFlowManager.WorkFlow();
            }
            catch (Exception e)
            {
                string message = "Otp AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }
    }
}
