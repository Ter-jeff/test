using Automation.GenerateIgxl.Basic;
using Automation.GenerateIgxl.BinCut;
using Automation.GenerateIgxl.EFuse;
using Automation.GenerateIgxl.PreAction;
using Automation.Static;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.Project
{
    [TestClass]
    public class Hidra : TestBase
    {
        [TestMethod]
        public void Lite_Hidra()
        {
            string subName = "hidra";
            RunHidraTest(subName, (testPlans, scghSheets) =>
            {
                using (var preActionMain = new PreActionMain())
                {
                    preActionMain.Execute(null);
                }

                using (var basicMain = new BasicMain())
                {
                    basicMain.Execute(null);
                }

                using (var eFuseMain = new EFuseMain())
                {
                    eFuseMain.Execute(null);
                }


                using (var binCutMain = new BinCutMain())
                {
                    binCutMain.GenerateBincutRelatedFiles();
                    binCutMain.Execute(null);
                }

                TestProgram.Print();
            });
        }
    }
}
