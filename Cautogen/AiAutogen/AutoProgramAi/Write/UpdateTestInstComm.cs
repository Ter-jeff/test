using System;
using System.Collections.Generic;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlSheets;

namespace Cautogen.AiAutogen.AutoProgramAi.Write
{
    public class UpdateTestInstComm
    {
        private bool _isCSharp = false;
        public UpdateTestInstComm(bool isCSharp)
        {
            _isCSharp = isCSharp;
        }

        public InstanceSheet Work(List<InstanceSheet> instSheets, VbtFunctionLib vbtFunctionLib)
        {
            var commInstSheet = instSheets.FirstOrDefault(x => x.Name.Equals("TestInst_Common", StringComparison.OrdinalIgnoreCase));
            var allInstName = instSheets.SelectMany(x => x.Rows).Where(x => !x.IsBackup).Select(x => x.TestName.ToUpper()).Distinct().ToList();

            WriteEnableCharzMode(commInstSheet, allInstName, vbtFunctionLib);
            return commInstSheet;
        }

        private void WriteEnableCharzMode(InstanceSheet modifiedSheet, List<string> allInstName, VbtFunctionLib vbtFunctionLib)
        {
            var instName = "Enable_Charz_mode";
            if (allInstName.Contains(instName.ToUpper()))
                return;

            if (!_isCSharp)
            {
                var newInstance = new InstanceRow();
                var vbtFunction = vbtFunctionLib.GetFunctionByName("TPmodeCharOn");
                newInstance.TestName = instName;
                newInstance.VbtType = "VBT";
                newInstance.VbtName = vbtFunction.FunctionName;
                newInstance.ArgList = vbtFunction.Parameters;
                modifiedSheet.AddRow(newInstance);

            }
            else
            {
                var newInstance = new InstanceRow();
                var vbtFunction = vbtFunctionLib.GetFunctionByName("TPmodeCharOn");
                newInstance.TestName = instName;
                newInstance.VbtType = ".NET";
                newInstance.VbtName = vbtFunction.FunctionName;
                newInstance.ArgList = vbtFunction.Parameters;
                modifiedSheet.AddRow(newInstance);

                //var newInstance = new InstanceRow { TestName = instName, Name = "CoreTestLibrary.Char.FunctionalTestCharMain.TPmodeCharOn", Type = ".NET" };
                //modifiedSheet.AddRow(newInstance);

                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = ConstData.EnableCzMode, VbtName = ConstData.TpModeOnModuleCSharp, VbtType = ".NET" });
                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = ConstData.DisableCzMode, VbtName = ConstData.TpModeOffModuleCSharp, VbtType = ".NET" });

                ////Add test instance to common test instance sheet

                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Hole_Ratio_Within_Spec", VbtName = "CoreTestLibrary.Char.FunctionalTestCharMain.CheckCharErrorCount", VbtType = ".NET", ArgList = "shmooAbnormalType,shmooAbnormalRatio_Hilimt", Args = new List<string> { "shmoo_hole", "0.1" } });
                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Allfail_Ratio_Within_Spec", VbtName = "CoreTestLibrary.Char.FunctionalTestCharMain.CheckCharErrorCount", VbtType = ".NET", ArgList = "shmooAbnormalType,shmooAbnormalRatio_Hilimt", Args = new List<string> { "all_fail", "0.1" } });
                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Check_Shmoo_Alarm_Ratio_Within_Spec", VbtName = "CoreTestLibrary.Char.FunctionalTestCharMain.CheckCharErrorCount", VbtType = ".NET", ArgList = "shmooAbnormalType,shmooAbnormalRatio_Hilimt", Args = new List<string> { "alarm", "0.1" } });
                ////_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Disable_Shmoo_Abnormal_Counter", VbtName = "DisableShmooAbnormalCounter", VbtType = ".NET" });
                //_CheckThenAdd(commonCharInstSheet, new InstanceRow { TestName = "Char_Setup_Gating", VbtName = "CoreTestLibrary.Char.FunctionalTestCharMain.CharSetupGating", VbtType = ".NET" });
            }
        }
    }
}
