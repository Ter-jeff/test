using System;
using System.Collections.Generic;
using System.Linq;

namespace Cautogen.AutoCZ.CharPostProcessor.Utility.VbtModuleManager
{
    public class VbtFunctionLib
    {
        /* properties */
        public const string FunctionalName = "functional_t_updated";
        public const string FunctionalCharName = "Functional_T_char";
        public const string VifName = "meas_freqvoltcurr_universal_func";
        public const string VirName = "meas_vir_io_universal_func";
        public const string DsscSearchName = "dssc_search";
        public const string FunctionalSpiName = "functional_t_updated_spi";
        public const string IdsVbtName = "dcvs_ids_main_current";
        public const string PpmuContinuity = "ppmu_continuity";
        public const string Uvi80Continuity = "UVI80_Continuity";
        public const string P2PShortPowerFvmi = "p2p_short_Power_FVMI";
        public const string VdiffFunc = "meas_vdiff_func";
        public const string DcvsPowerUpParallel = "dcvs_powerup_parallel ";
        public const string FreqSynMeasFreqCurr = "freqsyn_measfreqcurr_func";
        public const string MeasVohl = "meas_vohl_univeral_func_parallel";
        public const string DdrLpBkFunc2 = "opt_ddrlpbkfunc2";
        public const string PowerUp = "DCVS_PowerUp_Parallel";

        public static List<List<string>> ParamMappingList;
        public List<VbtFunction> VbtLib { get; set; }
        private static List<string> MissingParams { get; set; }

        /* constructor */
        public VbtFunctionLib()
        {
            VbtLib = new List<VbtFunction>();
            MissingParams = new List<string>();
            ParamMappingList = new List<List<string>>();
        }

        public bool IsFunctionExist(string functionName)
        {
            return VbtLib.Any(a => string.Equals(a.FunctionName, functionName, StringComparison.CurrentCultureIgnoreCase));
        }

        public VbtFunction GetFunctionByName(string functionName)
        {
            var vbtFunc = new VbtFunction();
            VbtFunction resultVbt = VbtLib.Find(a => string.Equals(a.FunctionName, functionName, StringComparison.CurrentCultureIgnoreCase)) ??
                            new VbtFunction();
            vbtFunc.FileName = resultVbt.FunctionName;
            vbtFunc.FunctionName = functionName;
            vbtFunc.Parameters = resultVbt.Parameters;
            vbtFunc.ArgList = new List<string>();
            vbtFunc.NameSpace = resultVbt.NameSpace;
            vbtFunc.Type = resultVbt.Type;
            foreach (string arg in resultVbt.ArgList)
            {
                vbtFunc.ArgList.Add(arg);
            }

            return vbtFunc;
        }

        public void AddVbtFunction(VbtFunction vbtFunction)
        {
            VbtLib.Add(vbtFunction);
        }

        public void AddVbtFunctionRange(List<VbtFunction> vbtFunction)
        {
            VbtLib.AddRange(vbtFunction);
            //SetGlobalFunctionName();
        }

        public static void CheckMissingParamter(string functionName, string paramName)
        {
            if (MissingParams.Contains(functionName + "&" + paramName))
            {
                return;
            }

            MissingParams.Add(functionName + "&" + paramName);
        }
    }
}
