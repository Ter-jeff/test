using System;

using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Reader.ConfigFile.NamingRule.Base;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

namespace Automation.GenerateIgxl.EVS
{
    public class EvsMain : WorkFlowBase<ParaData>
    {
        private EvsInputData _evsInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                _evsInputData = new EvsInputManager(EpWorkbook.TestPlanWorkbook).Read();

                return true;

            }
            catch (Exception e)
            {
                Response.Report("Evs has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                ScanConfig config = SettingStatic.ScanConfig;
                Response.Report("Generating Evs Instance sheet ...", percentage: 5);
                var evsInstanceFlow = new EvsInstanceMain(config, _evsInputData.EvsInstanceSheets);
                evsInstanceFlow.WorkFlow();
                Response.Report("Evs Completed!", percentage: 100);
            }
            catch (Exception e)
            {
                string message = "Evs AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }
    }
}
