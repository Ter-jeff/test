using System.Collections.Generic;
using System.IO;

using Automation.Const;

using CommonLib.Extension;

using ECIDSortAutogen;

using FileDiffLib;

using IgxlLib;
using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using OpCode = IgxlLib.IgxlConst.OpCode;

namespace Automation.Test.UT.PostAction
{
    [TestClass]
    public class IgxlTests : FunctionTestBase
    {
        private static IgxlLoader _igxlDataLoader = null!;

        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            string igxl = Path.Combine(InputPath, "PostAction", "komodo.igxl");
            var types = new List<EnumSheetType> { EnumSheetType.DTFlowtableSheet, EnumSheetType.DTBintablesSheet, EnumSheetType.DTTestInstancesSheet };
            _igxlDataLoader = new IgxlLoader(igxl, types);
        }

        [TestMethod]
        [TestCategory("ExcludeFromMutationTest")]
        public void EcidProgramGenerateMainTest()
        {
            string subName = "EcidProgramGenerateMain";
            string outputPath = Path.Combine(OutputPath, "PostAction", subName);
            string expectPath = Path.Combine(ExpectPath, "PostAction", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            SubFlowSheet? mainInitEnableWdSheet = _igxlDataLoader.FlowSheets.Find(x => x.Name.EqualsIgnoreCase(IgxlWorkBook.FlowTableMainInitEnableWd));
            SubFlowSheet? ecidSubFlowSheet = _igxlDataLoader.FlowSheets.Find(x => x.Name.EqualsIgnoreCase(EFuseConst.EcidFlowSheet));
            SubFlowSheet? ecidDeidSubFlowSheet = _igxlDataLoader.FlowSheets.Find(x => x.Name.EqualsIgnoreCase(EFuseConst.EcidDeidFlowSheet));
            InstanceSheet? ecidInstanceSheet = _igxlDataLoader.InstanceSheets.Find(x => x.Name.EqualsIgnoreCase(EFuseConst.TestInstSheet));

            var ecidProgramGenerateMain = new EcidProgramGenerateMain();
            if (!ecidSubFlowSheet!.Rows.Count.Equals(0))
            {
                EcidProgramGenerateMain.GenSyntaxCheckFlow(ecidSubFlowSheet);
                EcidProgramGenerateMain.GenEcidSortingFlow(ecidSubFlowSheet);
            }
            if (!ecidDeidSubFlowSheet!.Rows.Count.Equals(0))
            {
                EcidProgramGenerateMain.GenSyntaxCheckFlow(ecidDeidSubFlowSheet);
                EcidProgramGenerateMain.GenEcidSortingFlow(ecidDeidSubFlowSheet);
            }
            if (ecidInstanceSheet != null)
            {
                EcidProgramGenerateMain.GenEcidSortingInst(ecidInstanceSheet);
                ecidInstanceSheet.Write(Path.Combine(outputPath, ecidInstanceSheet.Name + ".txt"));
            }
            EcidProgramGenerateMain.GenEnableToMainInitEnableWd(mainInitEnableWdSheet!, "ECIDSort_Enable", OpCode.DisableFlowWd, EFuseConst.Efuse);
            EcidProgramGenerateMain.GenEnableToMainInitEnableWd(mainInitEnableWdSheet!, "Enable_ECID_Sorting_2CMode", OpCode.DisableFlowWd, EFuseConst.Efuse);

            mainInitEnableWdSheet!.Write(Path.Combine(outputPath, mainInitEnableWdSheet.Name + ".txt"));
            ecidSubFlowSheet.Write(Path.Combine(outputPath, ecidSubFlowSheet.Name + ".txt"));
            ecidDeidSubFlowSheet.Write(Path.Combine(outputPath, ecidDeidSubFlowSheet.Name + ".txt"));

            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
