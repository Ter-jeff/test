using System;

using Automation;
using Automation.GenerateIgxl;
using Automation.InputManager;
using Automation.InputManager.Data;
using Automation.PreCheck.AllParaData;
using Automation.PreCheck.PreCheckManager;
using Automation.Static;

using CommonLib.Enums;

using LogLib.Static;

namespace RfLib.Basic
{
    public class BasicMainRf : WorkFlowBase<ParaData>
    {
        protected BasicInputData BasicInputData { get; set; } = new();
        private PreCheckManager<ParaData> _preCheckManager = null!;

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

        public override void WorkFlow(ParaData paraData)
        {
            try
            {
                var autoGen = new AutogenMainRf();

                autoGen.Execute(paraData);
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
