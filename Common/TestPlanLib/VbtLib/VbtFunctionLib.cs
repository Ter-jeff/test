using System.Collections.Generic;

using CommonLib.Extension;

using TestPlanLib.Static;

namespace TestPlanLib.VbtLib
{
    public class VbtFunctionLib
    {
        public List<Function> VbtLib { get; set; }

        public VbtFunctionLib()
        {
            VbtLib = [];
            VbtFunctionLibShared.GeneratedVbtFunctionDic.Clear();
        }

        public virtual Function GetFunctionByName(string functionName, string block, bool isCsharp = false)
        {
            var function = new Function();
            Function? result = VbtLib.Find(a => a.FunctionName.EqualsIgnoreCase(functionName) && a.Type == ".NET")
                              ?? VbtLib.Find(a => a.FunctionName.EqualsIgnoreCase(functionName));

            if (result == null)
            {
                VbtFunctionLibShared.ReportMissingLibraryModuleError(block, functionName, isCsharp);
                function.FunctionName = functionName;
                function.Block = block;
                function.IsFound = false;
                return function;
            }

            function.FunctionName = result.FunctionName;
            function.FileName = result.FileName;
            function.NameSpace = result.NameSpace;
            function.Type = result.Type;
            function.Parameters = result.Parameters;
            function.ParameterDefaults = result.ParameterDefaults;
            function.PatternDic = result.PatternDic;
            function.Block = block;
            function.IsFound = true;
            function.IsInterPose = result.IsInterPose;
            return function;
        }

        public void AddVbtFunction(Function function)
        {
            VbtLib.Add(function);
            SetGlobalFunctionName();
        }

        public void AddVbtFunctionRange(List<Function> functions)
        {
            VbtLib.AddRange(functions);
            SetGlobalFunctionName();
        }

        public void Clear()
        {
            VbtLib.Clear();
        }

        private void SetGlobalFunctionName()
        {
            VbtFunctionLibShared.Ids = VbtLib.Exists(a => a.FunctionName.EqualsIgnoreCase("IDSCurrent") && a.Type == ".NET")
                ? "IDSCurrent"
                : "ids_main_current";

            VbtFunctionLibShared.VifName = VbtLib.Exists(a => a.FunctionName.EqualsIgnoreCase("HardIPUniversalFunction") && a.Type == ".NET")
                ? "HardIPUniversalFunction"
                : "meas_freqvoltcurr_universal_func";

            VbtFunctionLibShared.FunctionalName = VbtLib.Exists(a => a.FunctionName.EqualsIgnoreCase("FuncTestMain") && a.Type == ".NET")
                ? "FuncTestMain"
                : "functional_t_updated";

            VbtFunctionLibShared.PowerUp = VbtLib.Exists(a => a.FunctionName.EqualsIgnoreCase("PowerUp") && a.Type == ".NET")
                ? "PowerUp"
                : "PowerUp_Parallel";

            VbtFunctionLibShared.PowerDown = VbtLib.Exists(a => a.FunctionName.EqualsIgnoreCase("PowerDown") && a.Type == ".NET")
                ? "PowerDown"
                : "PowerDown_Parallel";

            VbtFunctionLibShared.HardIpmtdTest = VbtLib.Exists(a => a.FunctionName.EqualsIgnoreCase("HardIPMultiTimeDomain") && a.Type == ".NET")
                ? "HardIPMultiTimeDomain"
                : "hardip_mtd_test";
        }
    }
}
