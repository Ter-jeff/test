using System;
using System.Linq;

using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

namespace Automation.GenerateIgxl.HTOL
{
    public class HtolMain : WorkFlowBase<ParaData>
    {
        private HtolInputData _htolInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                _htolInputData = new HtolInputManager(EpWorkbook.TestPlanWorkbook).Read();

                return true;
            }
            catch (Exception e)
            {
                Response.Report("HTOL Action has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                if (TestPlanStatic.HtolInstanceSheets != null && TestPlanStatic.HtolInstanceSheets.Any())
                {
                    var htolInstanceMain = new HtolInstanceMain(SettingStatic.ScanConfig, _htolInputData.BinCutInstanceSheets);
                    Response.Report("Generating Non BinCut Instance sheet ...", percentage: 75);
                    htolInstanceMain.WorkFlow();
                }

                Response.Report("HTOL Completed!", percentage: 100);
            }
            catch (Exception e)
            {
                string message = "HTOL AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }
    }
}
