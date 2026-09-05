using System.Collections.Generic;
using System.IO;

using Automation.GenerateIgxl.HardIp.AutoGenBusiness;
using Automation.Static;

using FileDiffLib;

using IgxlLib.IgxlBase;
using IgxlLib.IgxlBase.MultiRow;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.Test.UT.HardIp
{
    [TestClass]
    public class TmpsGeneratorTests : FunctionTestBase
    {
        [TestMethod]
        public void TmpsGeneratorTest()
        {
            string subName = "TmpsGenerator";
            string outputPath = Path.Combine(OutputPath, "HardIp", subName);
            string expectPath = Path.Combine(ExpectPath, "HardIp", subName);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            _ = Directory.CreateDirectory(outputPath);

            // Arrange
            LocalSpecs.TarFolder = outputPath;
            LocalSpecs.Options.AdpativeCoolingBlock = "A,IDS,C";
            LocalSpecs.Options.AdpativeCoolingSuBlock = "D,E,F";
            Dictionary<string, SubFlowSheet> subFlowSheets = new Dictionary<string, SubFlowSheet>
            {
                {
                    "Flow_Hardip_A", new SubFlowSheet("")
                    {
                        Name = "Flow_Hardip_A" ,
                        Rows = new FlowRows
                        {
                            new FlowRow { Job = "CP2" , Opcode = "Test" , Parameter = "A_D_NV" , HiLim="100" , Units = "C"},
                            new FlowRow { Opcode = "Test" , Parameter = "B_E_NV"}
                        }
                    }

                },
                {
                    "DCTEST_IDS", new SubFlowSheet("")
                    {
                        Name = "DCTEST_IDS" ,
                        Rows = new FlowRows
                        {
                            new FlowRow { Opcode = "Test" , Parameter = "IDS_E_NV"}
                        }
                    }

                }
            };
            // Act
            new TmpsGenerator().GenTmps(subFlowSheets);
            TestProgram.Print();

            // Assert
            bool fail = new FileComparisonReport(subName).IsFail(outputPath, expectPath, true);
            if (fail)
            {
                Assert.Fail("Unit Test Fail!!!");
            }
        }
    }
}
