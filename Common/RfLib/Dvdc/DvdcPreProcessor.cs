using Automation.GenerateIgxl.HardIp.HardIpPreCheck;
using Automation.GenerateIgxl.HardIp.HardIPUtility;
using Automation.GenerateIgxl.HardIp.InputReader;
using Automation.InputManager.Data;

using LogLib.Static;

namespace RfLib.Dvdc
{
    internal static class DvdcPreProcessor
    {
        internal static void PreProcess(HardIpInputData hardIpInputData)
        {
            #region Add Cap-Src manager for register assignment. Added on 2016/4/21
            if (hardIpInputData.PlanDic.Count > 0)
            {
                var regAssignParser = new RegisterAssignParser();
                regAssignParser.ParseRegisterAssign(hardIpInputData.PlanDic);
            }
            #endregion

            Response.Report("Cross check the input file", percentage: 20);
            var preCheckMain = new HardIpPreCheckMain();
            preCheckMain.Check(hardIpInputData.PlanDic, hardIpInputData.HardIpDcSheet);

            #region divide pattern according to pattern.ForceCondition
            var parseTestPlanMain = new ParseTestPlanByCondition();
            parseTestPlanMain.ParseTestPlanPatternByCondition(hardIpInputData.PlanDic);
            #endregion

            #region Pre-Process for Misc Info
            MiscInfoPreProcessor.PreProcessMiscInfo(hardIpInputData.PlanDic);
            #endregion
        }
    }
}
