using System;

using Automation.GenerateIgxl.EFuse.Business;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.PreCheck.PreCheckManager;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

namespace Automation.GenerateIgxl.EFuse
{
    public class EFuseMain : WorkFlowBase<ParaData>
    {
        protected EFuseInputData EFuseInputData;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                EFuseInputData = new EFuseInputManager(EpWorkbook.TestPlanWorkbook).Read();

                if (!LocalSpecs.Options.BypassPreCheck)
                {
                    new EfusePreCheckManager().Check(EFuseInputData);
                }

                return true;
            }
            catch (Exception e)
            {
                Response.Report("Efuse has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                var efuseWorkFlowManager = new EfuseGenerate(EFuseInputData);
                efuseWorkFlowManager.WorkFlow();
            }
            catch (Exception e)
            {
                string message = "Efuse AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }
    }
}
