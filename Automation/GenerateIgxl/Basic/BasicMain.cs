using System;

using Automation.GenerateIgxl.Basic.Business;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.PreCheck.PreCheckManager;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

namespace Automation.GenerateIgxl.Basic
{
    public class BasicMain : WorkFlowBase<ParaData>
    {
        protected BasicInputData BasicInputData { get; set; }
        private PreCheckManager<ParaData> _preCheckManager;

        public override bool PreCheckFlow(ParaData paraData)
        {
            try
            {
                BasicInputData = new BasicInputManager(EpWorkbook.TestPlanWorkbook).Read();
                if (!LocalSpecs.Options.BypassPreCheck)
                {
                    _preCheckManager = new BasicPreCheckManager(EpWorkbook.TestPlanWorkbook, paraData, BasicInputData);
                    _preCheckManager.PreCheckAll();
                }

                return true;
            }
            catch (Exception e)
            {
                Response.Report("Basic has errors : " + e.StackTrace, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
                return false;
            }
        }

        public override void WorkFlow(ParaData para)
        {
            try
            {
                var autoGen = new AutogenMainAp
                {
                    BasicInputData = BasicInputData
                };

                autoGen.WorkFlow(para);

                autoGen.HarvestCheck();
            }
            catch (Exception e)
            {
                string message = "Basic AutoGen Failed: " + e.StackTrace;
                Response.Report(message, EnumMessageLevel.Error, 0);
                GenerateIgxlMain.ReturnValue = 1;
            }
        }
    }
}
