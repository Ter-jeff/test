using System;
using System.Linq;

using Cautogen.AutoCZ.CharPostProcessor.Bussiness;
using Cautogen.AutoCZ.CharPostProcessor.IGLinkProcessor.DataStructure;
using Cautogen.AutoCZ.CharPostProcessor.LocalSpec;
using Cautogen.AutoCZ.CharPostProcessor.Utility.UtilityFunctions;

using IgxlLib.IgxlSheets;

using LogLib.Utility;

namespace Cautogen.AutoCZ.CharPostProcessor.Controller
{
    public class GeneratorMain
    {
        //_param, _param.JobName, _param.GenFlowProdFlow, _param.GenTestInstCharPlan,
        //_param.GenInstancePatList, _param.PatternFolder, _param.HardIpInfoAllDict, out enableWord,
        //            _param.GenTNum, _param.GenTmpsOnFlow, _param.FlowTmpsName, _param.GenPmode, _param.GenPatSub
        public static void Run(InputParam inputParam, out string enableWord)
        {
            try
            {
                //var log = LogManager.GetCurrentClassLogger();

                LogHelper.Info("Generating AC Spec ...");
                var acSpecGenerator = new AcSpecGenerator(inputParam);
                acSpecGenerator.Generate(LocalSpecs.CharPlanSheets);
                acSpecGenerator.GenerateByTimeSettingsSheet(LocalSpecs.CharPlanSheets);

                //inst sheets
                LogHelper.Info("Generating Instance ...");
                new InstanceGenerator(inputParam).Generate(LocalSpecs.CharPlanSheets);

                //shmoo setup sheet
                LogHelper.Info("Generating Shmoo ...");
                GeneralFunc.WriteMessage("Generating shmoo setup sheet... ");
                new ShmooGenerator(inputParam.GenCSharp).Generate();

                //flow sheets
                LogHelper.Info("Generating Flow ...");
                var flowGenerator = new FlowGenerator(inputParam);
                SubFlowSheet flowChar = flowGenerator.Generate(LocalSpecs.CharPlanSheets);
                enableWord = string.Join(",",
                    flowChar.Rows.Select(x => x.Enable).Where(y => !string.IsNullOrEmpty(y)).Distinct());

                //bintable sheets
                LogHelper.Info("Generating BinTable ...");
                BinTableGenerator.Generate(LocalSpecs.CharPlanSheets);

                //GlobalSpec Sheet
                LogHelper.Info("Generating Global Spec ...");
                GlobalSpecGenerator.Generate();

                //PatSets_All_CZ, Pattern_Subroutine
                LogHelper.Info("Generating PatSet ...");
                var patSetGenerator = new PatSetGenerator(inputParam);
                patSetGenerator.Generate();
                if (inputParam.UseNewTChar)
                {
                    patSetGenerator.GeneratePatSetsCz(LocalSpecs.CharPlanSheets);
                }

                //DFC
                if (LocalSpecs.DFCSheet != null)
                {
                    DfcSheetGenerator.Generate();
                }

                OtherSheetGenerator.Generate();
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.StackTrace);
                LogHelper.Error(ex.Message);
                throw new Exception("Generation failed! " + ex.Message);
            }
        }
    }
}
